"""
SACEMRepertoire — Scraper du répertoire public SACEM
=====================================================
Utilisation depuis un script externe :

    from sacem_repertoire import SACEMRepertoire

    sacem = SACEMRepertoire()
    resultats = sacem.rechercher("mohombi")
    sacem.export_csv(resultats, "C:/exports")

Utilisation en ligne de commande (piloté depuis VB) :

    python sacem_repertoire.py --query "mohombi" --filtre parties --csv "C:/exports"
    python sacem_repertoire.py --query "mohombi" --ipi 00123456789 00987654321

Sortie stdout JSON ligne par ligne (pour VB) :
    {"type": "pagination", "total": 147, "max_page": 3}
    {"type": "oeuvres", "page": 1, "oeuvres": [...]}
    {"type": "done", "total_recupere": 147, "total_filtre_ipi": 12}
    {"type": "error", "message": "..."}
"""

import re
import ast
import csv
import json
import time
import logging
import argparse
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import requests

# ─── Logging (stderr uniquement — stdout réservé au JSON) ────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[logging.StreamHandler(__import__("sys").stderr)],
)
log = logging.getLogger("SACEM")


# ─── Dataclasses ─────────────────────────────────────────────────────────────

@dataclass
class Adresse:
    adresse1:   str = ""
    adresse2:   str = ""
    cp:         str = ""
    ville:      str = ""
    pays_code:  str = ""
    pays_nom:   str = ""
    telephones: list[str] = field(default_factory=list)

    @property
    def adresse_complete(self) -> str:
        parts = [self.adresse1, self.adresse2, f"{self.cp} {self.ville}".strip(), self.pays_nom]
        return ", ".join(p for p in parts if p)

    @property
    def telephone_principal(self) -> str:
        return self.telephones[0] if self.telephones else ""


@dataclass
class Partie:
    """Auteur, compositeur ou éditeur (workInterestedParties)."""
    nom:            str = ""
    prenom:         str = ""
    raison_sociale: str = ""
    ipi:            Optional[str] = None
    roles:          list[str] = field(default_factory=list)
    mail:           str = ""
    adresse:        Optional[Adresse] = None

    @property
    def nom_complet(self) -> str:
        return self.raison_sociale or " ".join(p for p in [self.prenom, self.nom] if p)

    @property
    def roles_str(self) -> str:
        return " / ".join(self.roles)


@dataclass
class Interprete:
    """Interprète (workPerformers)."""
    nom:            str = ""
    prenom:         str = ""
    raison_sociale: str = ""
    ipi:            Optional[str] = None
    roles:          list[str] = field(default_factory=list)
    adresse:        Optional[Adresse] = None

    @property
    def nom_complet(self) -> str:
        return self.raison_sociale or " ".join(p for p in [self.prenom, self.nom] if p)

    @property
    def roles_str(self) -> str:
        return " / ".join(self.roles)


@dataclass
class Oeuvre:
    token:          str
    titre:          str
    sous_titres:    list[str]
    iswc:           str
    genre_code:     str
    genre_label:    str
    duree_secondes: int
    type_oeuvre:    Optional[str]
    parties:        list[Partie]     = field(default_factory=list)
    interpretes:    list[Interprete] = field(default_factory=list)
    ipi_match:      bool             = False  # True si matche le filtre IPI

    CODES_AUTEUR      = {"A", "AE", "AD"}
    CODES_COMPOSITEUR = {"C", "CE", "AR"}
    CODES_AC          = {"CA", "A+C", "AC"}   # les deux à la fois
    CODES_EDITEUR     = {"E"}
    CODES_SOUS_ED     = {"SE"}
    CODES_REALISATEUR = {"R", "RE", "ES"}      # R = code réel SACEM (Réalisateur)

    @property
    def auteurs(self) -> list["Partie"]:
        """Personnes physiques avec rôle auteur (A, AE, AD) ou double (CA/A+C)."""
        return [p for p in self.parties
                if not p.raison_sociale
                and any(r.upper() in (self.CODES_AUTEUR | self.CODES_AC) for r in p.roles)]

    @property
    def compositeurs(self) -> list["Partie"]:
        """Personnes physiques avec rôle compositeur (C, CE, AR) ou double (CA/A+C)."""
        return [p for p in self.parties
                if not p.raison_sociale
                and any(r.upper() in (self.CODES_COMPOSITEUR | self.CODES_AC) for r in p.roles)]

    @property
    def editeurs(self) -> list["Partie"]:
        """Éditeurs principaux (rôle E) — raison sociale uniquement."""
        return [p for p in self.parties
                if p.raison_sociale
                and any(r.upper() in self.CODES_EDITEUR for r in p.roles)]

    @property
    def sous_editeurs(self) -> list["Partie"]:
        """Sous-éditeurs (rôle SE) — raison sociale uniquement."""
        return [p for p in self.parties
                if p.raison_sociale
                and any(r.upper() in self.CODES_SOUS_ED for r in p.roles)]

    @property
    def realisateurs(self) -> list["Partie"]:
        """Réalisateurs / sound engineers (rôle RE, ES)."""
        return [p for p in self.parties
                if any(r.upper() in self.CODES_REALISATEUR for r in p.roles)]

    @property
    def duree_str(self) -> str:
        if not self.duree_secondes:
            return ""
        m, s = divmod(self.duree_secondes, 60)
        return f"{m}:{s:02d}"

    def tous_les_ipi(self) -> set[str]:
        """Tous les IPI présents sur l'oeuvre (parties + interprètes)."""
        return {str(x.ipi).strip() for x in [*self.parties, *self.interpretes] if x.ipi}

    @staticmethod
    def _fmt_partie(p: "Partie") -> str:
        """IPI : NOM [rôles]"""
        parts = []
        if p.ipi:
            parts.append(f"{p.ipi} : {p.nom_complet}")
        else:
            parts.append(p.nom_complet)
        if p.roles:
            parts.append(f"[{p.roles_str}]")
        return " ".join(parts)

    @staticmethod
    def _fmt_editeur(e: "Partie") -> str:
        """IPI : Raison sociale [rôle] | adresse | tél"""
        if e.ipi:
            info = f"{e.ipi} : {e.nom_complet}"
        else:
            info = e.nom_complet
        if e.roles:
            info += f" [{e.roles_str}]"
        extras = []
        if e.adresse:
            extras.append(e.adresse.adresse_complete)
            if e.adresse.telephone_principal:
                extras.append(e.adresse.telephone_principal)
        if e.mail:
            extras.append(e.mail)
        if extras:
            info += " | " + " | ".join(ex for ex in extras if ex)
        return info

    def _fmt_ayants_droits(self) -> str:
        """
        Colonne synthèse phase 1 + phase 2.
        Regroupe par personne (clé = IPI si dispo, sinon nom_complet).
        Agrège les rôles. Ordre : INT > C+A > C > A > E > SE.
        Format ligne : "ROLES : NOM"  (ex: "INT,C,A : BOOBA")
        """
        ORDRE = {"INT": 0, "A+C": 1, "CA": 1, "AC": 1,
                 "C": 2, "CE": 2, "AR": 2,
                 "A": 3, "AE": 3, "AD": 3,
                 "E": 4, "SE": 5,
                 "R": 6, "RE": 6, "ES": 6}   # R = code réel SACEM

        # clé → (nom_complet, set de rôles, priorité min)
        personnes: dict[str, list] = {}

        def _ajouter(ipi, nom, roles):
            cle = str(ipi).strip() if ipi else nom.strip()
            if not cle:
                return
            if cle not in personnes:
                personnes[cle] = [nom, set(), 99]
            personnes[cle][1].update(r.upper() for r in roles)
            prio = min((ORDRE.get(r.upper(), 99) for r in roles), default=99)
            personnes[cle][2] = min(personnes[cle][2], prio)

        # Index nom → ipi depuis les interprètes (pour fusionner avec parties)
        interp_ipi_par_nom: dict[str, Optional[str]] = {}
        for i in self.interpretes:
            interp_ipi_par_nom[i.nom_complet.strip().upper()] = i.ipi

        for i in self.interpretes:
            _ajouter(i.ipi, i.nom_complet, i.roles or ["INT"])

        for p in self.parties:
            roles_clean = [r for r in p.roles if r.strip()]
            # Si cette personne est aussi interprète → ajouter INT à ses rôles
            nom_up = p.nom_complet.strip().upper()
            if nom_up in interp_ipi_par_nom:
                roles_clean = list(set(roles_clean + ["INT"]))
            _ajouter(p.ipi, p.nom_complet, roles_clean if roles_clean else (p.roles or []))

        # Trier par priorité
        triees = sorted(personnes.values(), key=lambda x: x[2])

        lignes = []
        for nom, roles, _ in triees:
            # Simplifier l'affichage des rôles
            roles_aff = []
            has_ca = bool(roles & {"A+C", "CA", "AC"})
            has_int = "INT" in roles
            if has_int:
                roles_aff.append("INT")
            if has_ca:
                roles_aff.append("A+C")
            else:
                if roles & {"C", "CE", "AR"}:
                    roles_aff.append("C")
                if roles & {"A", "AE", "AD"}:
                    roles_aff.append("A")
            if "E" in roles:
                roles_aff.append("E")
            if "SE" in roles:
                roles_aff.append("SE")
            if roles & {"R", "RE", "ES"}:
                roles_aff.append("RE")
            prefix = ",".join(roles_aff) if roles_aff else (",".join(sorted(roles)) or "?")
            lignes.append(f"{prefix} : {nom}")

        return "\n".join(lignes)

    def to_dict(self) -> dict:
        """Dict plat — chaque personne sur sa propre ligne (séparateur \\n)."""
        return {
            "titre":            self.titre,
            "sous_titres":      " | ".join(self.sous_titres),
            "iswc":             self.iswc,
            "genre":            self.genre_label,
            "duree":            self.duree_str,
            "type":             self.type_oeuvre or "",
            "token":            self.token,
            "ayants_droits":    self._fmt_ayants_droits(),
            "auteurs":          "\n".join(self._fmt_partie(p) for p in self.auteurs),
            "compositeurs":     "\n".join(self._fmt_partie(p) for p in self.compositeurs),
            "editeurs":         "\n".join(self._fmt_editeur(e) for e in self.editeurs),
            "sous_editeurs":    "\n".join(self._fmt_editeur(e) for e in self.sous_editeurs),
            "interpretes":      "\n".join(self._fmt_partie(i) for i in self.interpretes),
            "realisateurs":     "\n".join(self._fmt_partie(p) for p in self.realisateurs),
            "ipi_match":        self.ipi_match,
        }


@dataclass
class ResultatRecherche:
    query:      str
    filtre:     str
    total:      int        # total annoncé par SACEM
    max_page:   int
    ipi_filter: list[str]    = field(default_factory=list)
    oeuvres:    list[Oeuvre] = field(default_factory=list)

    # Sources brutes — alimentées pendant rechercher()
    _source_pages: dict[int, str] = field(default_factory=dict,  repr=False)

    @property
    def total_recupere(self) -> int:
        return len(self.oeuvres)

    @property
    def total_filtre_ipi(self) -> int:
        return sum(1 for o in self.oeuvres if o.ipi_match)


# ─── Classe principale ────────────────────────────────────────────────────────

class SACEMRepertoire:

    EL_PER_PAGE = 50  # valeur max fiable testée (au-delà le serveur bugue)

    _HEADERS_GET = {
        "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection":      "keep-alive",
        "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36",
    }

    _HEADERS_POST = {
        "Accept":          "*/*",
        "Accept-Language": "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection":      "keep-alive",
        "Content-Type":    "application/json",
        "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36",
    }

    def __init__(
        self,
        base_url:  str   = "https://repertoire.sacem.fr",
        delay:     float = 0.5,
        timeout:   int   = 15,
        max_pages: Optional[int] = None,
    ):
        self.base_url         = base_url.rstrip("/")
        self.search_url       = self.base_url + "/resultats"
        self.more_results_url = self.base_url + "/en/more-results"
        self.delay            = delay
        self.timeout          = timeout
        self.max_pages        = max_pages
        self._last_total      = 0
        self.session          = requests.Session()
        self.session.headers.update(self._HEADERS_GET)

    # ── API publique ──────────────────────────────────────────────────────────

    def rechercher(
        self,
        query:         str,
        filtre:        str = "parties",
        ipi_filter:    Optional[list[str]] = None,
        tokens_exclus: Optional[set] = None,
        sans_details:  bool = False,
        cles_exclues:  Optional[set] = None,
    ) -> ResultatRecherche:
        """
        Recherche complète via POST uniquement (plus de GET HTML initial).
        Émet les résultats sur stdout JSON ligne par ligne.
        tokens_exclus : tokens déjà enrichis par une requête précédente — skippés à l'enrichissement.
        cles_exclues  : clés ISWC/titre déjà connues (inter-requêtes) — filtrées avant emit pour éviter doublons.
        """
        ipi_set     = {str(i).strip() for i in ipi_filter} if ipi_filter else None
        excl_set    = tokens_exclus or set()
        vus         = set(cles_exclues) if cles_exclues else set()
        cles_inter  = frozenset(vus)   # clés reçues via stdin = requêtes précédentes
        inter_vus   = set()            # clés inter déjà comptées comme communes (évite double-comptage)
        nb_exclus   = 0   # oeuvres communes inter-requêtes (dans cles_inter)
        nb_intra    = 0   # doublons intra-requête (SACEM pagine mal)
        # "titles.parties" : query = "TITRE,CREATEUR" — SACEM accepte ce format
        log.info(f"Recherche query={query!r} filtre={filtre!r} ipi={ipi_set} exclus={len(excl_set)} cles_exclues={len(vus)}")

        def _cle(o: "Oeuvre") -> str:
            return o.iswc.strip() if o.iswc.strip() else o.titre.strip().upper()

        def _filtrer(oeuvres: list) -> list:
            """Retire les oeuvres déjà vues, distingue inter vs intra."""
            nonlocal nb_exclus, nb_intra
            result = []
            for o in oeuvres:
                c = _cle(o)
                if c in vus:
                    if c in cles_inter and c not in inter_vus:
                        nb_exclus += 1   # première occurrence commune inter-requêtes
                        inter_vus.add(c)
                        log.info(f"  [INTER] {o.titre!r} (cle={c!r})")
                    else:
                        nb_intra += 1    # doublon interne ou inter déjà compté
                        log.info(f"  [INTRA] {o.titre!r} (cle={c!r})")
                        SACEMRepertoire._emit("intra", titre=o.titre, cle=c)
                else:
                    vus.add(c)
                    result.append(o)
            return result

        # ── Page 1 : POST avec elPerPage=50 → total + maxPage fiables ────────
        oeuvres_p1, json_brut_p1, max_page = self._get_page_suivante(query, filtre, 1, get_max_page=True)
        total  = self._last_total
        max_page = max_page or 1   # sécurité si SACEM ne retourne pas maxPage
        limite = min(max_page, self.max_pages) if self.max_pages else max_page

        oeuvres_p1 = _filtrer(oeuvres_p1)
        self._marquer_ipi(oeuvres_p1, ipi_set)

        resultat = ResultatRecherche(
            query=query, filtre=filtre,
            total=total, max_page=max_page,
            ipi_filter=list(ipi_set) if ipi_set else [],
        )
        resultat.oeuvres.extend(oeuvres_p1)
        resultat._source_pages[1] = json_brut_p1

        self._emit("pagination", total=total, max_page=limite)
        self._emit("oeuvres", page=1, oeuvres=[o.to_dict() for o in oeuvres_p1])

        # ── Pages 2..maxPage ─────────────────────────────────────────────────
        for page in range(2, limite + 1):
            log.info(f"  Page {page}/{limite}...")
            time.sleep(self.delay)
            oeuvres_page, json_brut, _ = self._get_page_suivante(query, filtre, page)
            oeuvres_page = _filtrer(oeuvres_page)
            self._marquer_ipi(oeuvres_page, ipi_set)
            resultat.oeuvres.extend(oeuvres_page)
            resultat._source_pages[page] = json_brut
            self._emit("oeuvres", page=page, oeuvres=[o.to_dict() for o in oeuvres_page])

        self._emit("done", total_recupere=resultat.total_recupere, total_filtre_ipi=resultat.total_filtre_ipi, exclus=nb_exclus, intra=nb_intra)
        log.info(f"Terminé. {resultat.total_recupere}/{total} oeuvres, {resultat.total_filtre_ipi} IPI match.")

        # ── Enrichissement IPI : GET détail pour chaque oeuvre ────────────────
        if sans_details:
            log.info("--sans-details : enrichissement skippé.")
            return resultat

        log.info("Enrichissement IPI via détail-oeuvre...")
        for i, oe in enumerate(resultat.oeuvres):
            if oe.token in excl_set:
                log.info(f"  [{i+1}/{resultat.total_recupere}] {oe.titre} — déjà enrichi, skippé")
                continue
            try:
                time.sleep(self.delay)
                detail = self._get_detail(oe.token, oe.titre, query, filtre)
                if detail:
                    oe.parties        = detail.parties
                    oe.interpretes    = detail.interpretes
                    oe.duree_secondes = detail.duree_secondes or oe.duree_secondes
                    self._marquer_ipi([oe], ipi_set)
                    self._emit("detail", index=i, oeuvre=oe.to_dict())
                    log.info(f"  [{i+1}/{resultat.total_recupere}] {oe.titre} OK")
            except Exception as e:
                log.warning(f"  [{i+1}] détail échoué pour {oe.token}: {e}")

        return resultat

    def export_csv(self, resultat: ResultatRecherche, dossier: str, seulement_ipi: bool = False) -> Path:
        """Crée le dossier, écrit le CSV + sources brutes."""
        out_dir = self._creer_dossier(resultat, dossier)
        oeuvres = [o for o in resultat.oeuvres if not seulement_ipi or o.ipi_match]

        if not oeuvres:
            log.warning("Aucune oeuvre à exporter.")
            return out_dir

        path = out_dir / f"{resultat.query}_{resultat.filtre}.csv"
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(oeuvres[0].to_dict().keys()), delimiter=";")
            writer.writeheader()
            for o in oeuvres:
                writer.writerow(o.to_dict())

        log.info(f"CSV : {path} ({len(oeuvres)} lignes)")
        self._ecrire_sources(resultat, out_dir / "sources")
        return out_dir

    def export_json(self, resultat: ResultatRecherche, dossier: str, seulement_ipi: bool = False) -> Path:
        """Écrit le JSON parsé dans le dossier de sortie."""
        out_dir = self._creer_dossier(resultat, dossier)
        oeuvres = [o for o in resultat.oeuvres if not seulement_ipi or o.ipi_match]

        path = out_dir / f"{resultat.query}_{resultat.filtre}.json"
        with path.open("w", encoding="utf-8") as f:
            json.dump({
                "query":            resultat.query,
                "filtre":           resultat.filtre,
                "total":            resultat.total,
                "total_recupere":   resultat.total_recupere,
                "total_filtre_ipi": resultat.total_filtre_ipi,
                "ipi_filter":       resultat.ipi_filter,
                "oeuvres":          [o.to_dict() for o in oeuvres],
            }, f, ensure_ascii=False, indent=2)

        log.info(f"JSON : {path}")
        self._ecrire_sources(resultat, out_dir / "sources")
        return out_dir

    # ── Requêtes ──────────────────────────────────────────────────────────────

    def stat_requete(self, query: str, filtre: str) -> int:
        """1 seul POST elPerPage=1 → retourne totalElements. Émet {"type":"stat",...}."""
        resp = self.session.post(
            self.more_results_url,
            json={"query": query, "filters": filtre, "page": 1, "elPerPage": 1},
            headers={
                **self._HEADERS_POST,
                "Referer": f"{self.search_url}?filters={filtre}&query={requests.utils.quote(query)}",
                "Origin":  self.base_url,
            },
            timeout=self.timeout,
        )
        resp.raise_for_status()
        total = resp.json().get("pagination", {}).get("totalElements", 0)
        self._emit("stat", query=query, filtre=filtre, total=total)
        return total

    def _get_page_suivante(self, query: str, filtre: str, page: int, get_max_page: bool = False) -> tuple:
        """POST /en/more-results → (oeuvres, json_brut, max_page)"""
        resp = self.session.post(
            self.more_results_url,
            json={"query": query, "filters": filtre, "page": page, "elPerPage": self.EL_PER_PAGE},
            headers={
                **self._HEADERS_POST,
                "Referer": f"{self.search_url}?filters={filtre}&query={requests.utils.quote(query)}",
                "Origin":  self.base_url,
            },
            timeout=self.timeout,
        )
        if resp.status_code != 200:
            raise ValueError(f"HTTP {resp.status_code} pour filtre={filtre!r} query={query!r}: {resp.text[:200]}")
        if not resp.text.strip():
            raise ValueError(f"Réponse vide pour filtre={filtre!r} query={query!r}")
        data = resp.json()
        pagination = data.get("pagination", {})
        self._last_total = pagination.get("totalElements", 0)
        oeuvres  = [self._parse_oeuvre(o) for o in data.get("paginatedData", [])]
        max_page = pagination.get("maxPage") if get_max_page else None
        return oeuvres, resp.text, max_page

    def _get_detail(self, token: str, titre: str, query: str, filtre: str) -> Optional[Oeuvre]:
        """
        GET /detail-oeuvre/{token}/{titre-urlencode} → Oeuvre avec IPI renseignés.
        Retourne None si la requête échoue ou si var data est introuvable.
        """
        titre_enc = requests.utils.quote(titre, safe="")
        url = f"{self.base_url}/detail-oeuvre/{requests.utils.quote(token, safe='')}/{titre_enc}"
        resp = self.session.get(
            url,
            params={"query": query, "filters": filtre},
            headers={
                **self._HEADERS_GET,
                "Referer": f"{self.search_url}?filters={filtre}&query={requests.utils.quote(query)}",
            },
            timeout=self.timeout,
        )
        resp.raise_for_status()
        match = re.search(r"var data\s*=\s*(\{.*?\});\s*(?:var|</script>)", resp.text, re.DOTALL)
        if not match:
            return None
        data = self._parse_js_object(match.group(1))
        return self._parse_oeuvre(data)

    # ── Parsers ───────────────────────────────────────────────────────────────

    def _parse_oeuvre(self, r: dict) -> Oeuvre:
        genre = r.get("workGenre") or {}
        return Oeuvre(
            token          = r.get("workToken", ""),
            titre          = r.get("workTitle", ""),
            sous_titres    = r.get("workSubtitles", []),
            iswc           = r.get("workISWCCode", ""),
            genre_code     = genre.get("workGenreCode", ""),
            genre_label    = genre.get("workGenreLabel", ""),
            duree_secondes = r.get("workDuration", 0) or 0,
            type_oeuvre    = r.get("workType"),
            parties        = [self._parse_partie(p)     for p in r.get("workInterestedParties", [])],
            interpretes    = [self._parse_interprete(p) for p in r.get("workPerformers", [])],
        )

    def _parse_partie(self, r: dict) -> Partie:
        return Partie(
            nom            = r.get("workInterestedPartyLastName", ""),
            prenom         = r.get("workInterestedPartyFirstName", ""),
            raison_sociale = r.get("workInterestedPartyCorporateName", ""),
            ipi            = r.get("workInterestedPartyIPICode"),
            roles          = [x.get("workInterestedPartyRoleCode", "") for x in r.get("workInterestedPartyRoles", [])],
            mail           = r.get("workInterestedPartyMail", ""),
            adresse        = self._parse_adresse(r.get("workInterestedPartyAddress")),
        )

    def _parse_interprete(self, r: dict) -> Interprete:
        return Interprete(
            nom            = r.get("workPerformerLastName", ""),
            prenom         = r.get("workPerformerFirstName", ""),
            raison_sociale = r.get("workPerformerCorporateName", ""),
            ipi            = r.get("workPerformerIPICode"),
            roles          = [x.get("workPerformerRoleCode", "") for x in r.get("workPerformerRoles", [])],
            adresse        = self._parse_adresse(r.get("workPerformerAddress")),
        )

    @staticmethod
    def _parse_adresse(r: Optional[dict]) -> Optional[Adresse]:
        if not r:
            return None
        pays = r.get("workInterestedPartyAddressCountry") or {}
        tels = r.get("workInterestedPartyPhoneNumber") or []
        return Adresse(
            adresse1   = r.get("workInterestedPartyAddress1") or "",
            adresse2   = r.get("workInterestedPartyAddress2") or "",
            cp         = r.get("workInterestedPartyAddressZipCode") or "",
            ville      = r.get("workInterestedPartyAddressCity") or "",
            pays_code  = pays.get("workInterestedPartyAddressCountryCode", ""),
            pays_nom   = pays.get("workInterestedPartyAddressCountryName", ""),
            telephones = [t["workInterestedPartyPhoneNumberValue"] for t in tels if t.get("workInterestedPartyPhoneNumberValue")],
        )

    @staticmethod
    def _parse_js_object(raw: str) -> dict:
        """Normalise JS single-quotes → dict Python via ast.literal_eval."""
        s = raw.replace(": null", ": None").replace(":null", ":None")
        s = re.sub(r'\btrue\b',  'True',  s)
        s = re.sub(r'\bfalse\b', 'False', s)
        return ast.literal_eval(s)

    # ── Export helpers ────────────────────────────────────────────────────────

    def _creer_dossier(self, resultat: ResultatRecherche, base: str) -> Path:
        """Crée et retourne le dossier de sortie (idempotent)."""
        nom = f"{resultat.query}_{resultat.filtre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        out = Path(base) / nom
        out.mkdir(parents=True, exist_ok=True)
        return out

    def _ecrire_sources(self, resultat: ResultatRecherche, sources_dir: Path):
        """Écrit les fichiers sources bruts dans sources_dir/."""
        sources_dir.mkdir(exist_ok=True)
        for page, contenu in resultat._source_pages.items():
            (sources_dir / f"page_{page:02d}.json").write_text(contenu, encoding="utf-8")

    @staticmethod
    def _marquer_ipi(oeuvres: list, ipi_set: Optional[set]):
        """Marque ipi_match=True sur les oeuvres dont un IPI est dans ipi_set."""
        if not ipi_set:
            return
        for o in oeuvres:
            o.ipi_match = bool(o.tous_les_ipi() & ipi_set)

    @staticmethod
    def _emit(type_: str, **kwargs):
        """Écrit une ligne JSON sur stdout (lu ligne par ligne par VB)."""
        print(json.dumps({"type": type_, **kwargs}, ensure_ascii=False), flush=True)


# ─── CLI ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    p = argparse.ArgumentParser(description="Scraper répertoire public SACEM",
                                fromfile_prefix_chars="@")
    p.add_argument("--query",         required=True)
    p.add_argument("--filtre",        default="parties", choices=["parties", "title", "titles,parties"])
    p.add_argument("--ipi",           nargs="*", default=None)
    p.add_argument("--base-url",      default="https://repertoire.sacem.fr")
    p.add_argument("--delay",         type=float, default=0.5)
    p.add_argument("--timeout",       type=int,   default=15)
    p.add_argument("--max-pages",     type=int,   default=None)
    p.add_argument("--csv",           default=None)
    p.add_argument("--json",          default=None)
    p.add_argument("--seulement-ipi", action="store_true")
    p.add_argument("--sans-details",      action="store_true")
    p.add_argument("--stat-seulement",    action="store_true")
    p.add_argument("--details-stdin",     action="store_true", help="Lire les TOKEN:TITRE:INDEX depuis stdin")
    p.add_argument("--details-seulement", nargs="*", default=None, metavar="TOKEN:TITRE:INDEX")
    p.add_argument("--exclusions-stdin",  action="store_true", help="Lire les clés ISWC/titre à exclure depuis stdin (une par ligne)")
    args = p.parse_args()

    sacem = SACEMRepertoire(
        base_url  = args.base_url,
        delay     = args.delay,
        timeout   = args.timeout,
        max_pages = args.max_pages,
    )

    try:
        # ── Mode stat seulement : 1 POST → totalElements, pas de pagination ──
        if args.stat_seulement:
            sacem.stat_requete(query=args.query, filtre=args.filtre)

        # ── Mode détails depuis stdin ──────────────────────────────────────
        elif args.details_stdin:
            ipi_set = set(args.ipi) if args.ipi else set()
            entrees = []
            for line in sys.stdin.read().strip().split("\n"):
                line = line.strip()
                if not line:
                    continue
                parts = line.split(":", 2)
                if len(parts) == 3:
                    tok, tit, idx_s = parts
                    try:
                        entrees.append((tok, tit, int(idx_s)))
                    except ValueError:
                        pass
            SACEMRepertoire._emit("pagination", total=len(entrees), max_page=1)
            for tok, tit, idx in entrees:
                try:
                    time.sleep(sacem.delay)
                    detail = sacem._get_detail(tok, tit, args.query, args.filtre)
                    if detail:
                        sacem._marquer_ipi([detail], ipi_set)
                        d = detail.to_dict()
                        has_ipi = any(": " in p for p in (d.get("auteurs", "") + "\n" + d.get("compositeurs", "")).split("\n") if p)
                        log.info(f"  [{idx}] {tit} OK — IPI: {has_ipi}")
                        SACEMRepertoire._emit("detail", index=idx, oeuvre=d)
                    else:
                        log.warning(f"  [{idx}] {tit} — None")
                except Exception as e:
                    log.warning(f"  [{idx}] échoué: {e}")
            SACEMRepertoire._emit("done", total_recupere=len(entrees), total_filtre_ipi=0)

        # ── Mode détails seulement (legacy @fichier) ───────────────────────
        elif args.details_seulement is not None:
            ipi_set = set(args.ipi) if args.ipi else set()
            entrees = []
            for item in args.details_seulement:
                parts = item.split(":", 2)
                if len(parts) == 3:
                    tok, tit, idx_s = parts
                    try:
                        entrees.append((tok, tit, int(idx_s)))
                    except ValueError:
                        pass
            SACEMRepertoire._emit("pagination", total=len(entrees), max_page=1)
            for tok, tit, idx in entrees:
                try:
                    time.sleep(sacem.delay)
                    detail = sacem._get_detail(tok, tit, args.query, args.filtre)
                    if detail:
                        sacem._marquer_ipi([detail], ipi_set)
                        d = detail.to_dict()
                        # Log debug IPI
                        has_ipi = any(': ' in p for p in (d.get('auteurs','') + '\n' + d.get('compositeurs','')).split('\n') if p)
                        log.info(f"  [{idx}] {tit} OK — IPI présents: {has_ipi}")
                        SACEMRepertoire._emit("detail", index=idx, oeuvre=d)
                    else:
                        log.warning(f"  [{idx}] {tit} — _get_detail retourne None")
                except Exception as e:
                    log.warning(f"  [{idx}] détail échoué pour {tok}: {e}")
            SACEMRepertoire._emit("done", total_recupere=len(entrees), total_filtre_ipi=0)
        else:
            cles_exclues: set = set()
            if args.exclusions_stdin:
                for line in sys.stdin.read().strip().split("\n"):
                    line = line.strip()
                    if line:
                        cles_exclues.add(line)
                log.info(f"Exclusions reçues via stdin : {len(cles_exclues)}")
            resultats = sacem.rechercher(query=args.query, filtre=args.filtre,
                                         ipi_filter=args.ipi,
                                         sans_details=args.sans_details,
                                         cles_exclues=cles_exclues if cles_exclues else None)
            if args.csv:
                sacem.export_csv(resultats,  args.csv,  seulement_ipi=args.seulement_ipi)
            if args.json:
                sacem.export_json(resultats, args.json, seulement_ipi=args.seulement_ipi)
    except Exception as e:
        SACEMRepertoire._emit("error", message=str(e))
        log.exception("Erreur fatale")
        sys.exit(1)
