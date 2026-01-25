# 📋 BALISES NON-SACEM - Documentation

## Vue d'ensemble

Ces balises permettent de gérer les **œuvres mixtes** (SACEM + NON-SACEM).

Elles sont générées automatiquement par `BalisesGenerator.vb` dans la méthode `GenerateNonSACEMBalises()`.

---

## 🔢 Balises de calcul (4)

| Balise | Description | Exemple |
|--------|-------------|---------|
| `[PartsSACEM]` | % total contrôlé par membres SACEM | `91,68` |
| `[PartsNonSACEM]` | % total contrôlé par membres NON-SACEM | `8,32` |
| `[PartsSACEM_Texte]` | % SACEM en toutes lettres | `quatre-vingt-onze virgule soixante-huit` |
| `[PartsNonSACEM_Texte]` | % NON-SACEM en toutes lettres | `huit virgule trente-deux` |

---

## 📝 Balises de listes SACEM (4)

| Balise | Description | Exemple |
|--------|-------------|---------|
| `[ListeAC_SACEM]` | Liste des AC membres SACEM | `SISSOKO Moussa, EMAM Bahi et MEZREB Sofiane` |
| `[ListeE_SACEM]` | Liste des E membres SACEM | `TH RECORDS, BELLUM PUBLISHING et SONY MUSIC PUBLISHING (France)` |
| `[ListeSACEM]` | Tous SACEM (AC puis E) | `SISSOKO Moussa, EMAM Bahi, TH RECORDS et BELLUM PUBLISHING` |
| `[ListeSACEM_Noms]` | Idem (pas de suffixe société pour SACEM) | Idem |

---

## 📝 Balises de listes NON-SACEM (4)

| Balise | Description | Exemple |
|--------|-------------|---------|
| `[ListeAC_NonSACEM]` | Liste des AC NON-SACEM avec société | `Oliver Zoega BJERG (KODA)` |
| `[ListeE_NonSACEM]` | Liste des E NON-SACEM avec société | `CUBEATZ MUSIC GMBH (GEMA)` |
| `[ListeNonSACEM]` | Tous NON-SACEM (AC puis E) avec société | `Oliver Zoega BJERG (KODA) et CUBEATZ MUSIC GMBH (GEMA)` |
| `[ListeNonSACEM_Noms]` | Tous NON-SACEM - noms seuls | `Oliver Zoega BJERG et CUBEATZ MUSIC GMBH` |

---

## 🔤 Balises pluriel/singulier (4)

| Balise | Singulier (1 NON-SACEM) | Pluriel (2+ NON-SACEM) |
|--------|-------------------------|------------------------|
| `[EstSont]` | `est` | `sont` |
| `[IlIls]` | `Il` | `Ils` |
| `[NestPasNeSontPas]` | `n'est pas` | `ne sont pas` |
| `[Pluriel]` | `` (vide) | `s` |

---

## 🎯 Balises spécifiques (2)

| Balise | Description | Exemple |
|--------|-------------|---------|
| `[NumPoint]` | Numéro du point dans Article 16 | `9` ou `10` |
| `[EditeurDe]` | Si E étranger édite AC SACEM | ` éditeur de MEZREB Sofiane` ou `` (vide) |

---

## 🔍 Balise de détection (1)

| Balise | Description | Valeurs |
|--------|-------------|---------|
| `[HasNonSACEM]` | Y a-t-il au moins un NON-SACEM ? | `True` ou `False` |

---

## 📄 Utilisation dans les templates

### template_paragrahs.docx

```
{START_MENTION_ART11_NONSACEM}
Il est ici précisé que l'ŒUVRE est coéditée avec d'une part, [ListeE_SACEM], et d'autre part [ListeNonSACEM], qui en contrôlent respectivement [PartsSACEM]% et [PartsNonSACEM]%.
Dans ce contexte, tous les pourcentages stipulés au présent contrat s'entendent au prorata de la part contrôlée par l'EDITEUR, soit [PartsSACEM]% de l'ŒUVRE.
{END_MENTION_ART11_NONSACEM}

{START_MENTION_ART16_NONSACEM}
[NumPoint]- Il est ici précisé que l'ŒUVRE est coéditée avec d'une part, [ListeE_SACEM], et d'autre part [ListeNonSACEM], qui en contrôlent respectivement [PartsSACEM]% et [PartsNonSACEM]%.
Dans ce contexte, tous les pourcentages stipulés au présent contrat s'entendent au prorata de la part contrôlée par l'EDITEUR, soit [PartsSACEM]% de l'ŒUVRE.
[ListeNonSACEM_Noms] [EstSont] membre[Pluriel] d'une autre société de gestion collective autre que la SACEM. [IlIls] [NestPasNeSontPas] signataire[Pluriel] du présent contrat mais de la lettre de répartition ci-après annexée.
{END_MENTION_ART16_NONSACEM}

{START_MENTION_COED_NONSACEM}
Il est ici précisé que l'ŒUVRE est coéditée avec d'une part, [ListeE_SACEM], et d'autre part [ListeNonSACEM], qui en contrôlent respectivement [PartsSACEM]% et [PartsNonSACEM]%.
Dans ce contexte, tous les pourcentages stipulés au présent contrat s'entendent au prorata de la part contrôlée par les COEDITEURS, soit [PartsSACEM]% de l'ŒUVRE.
[ListeNonSACEM_Noms] [EstSont] membre[Pluriel] d'une société de gestion collective autre que la SACEM. [IlIls] [NestPasNeSontPas] signataire[Pluriel] du présent contrat.
Les droits collectés par les Organismes de gestion collective ou indépendants seront répartis directement par ceux-ci.
{END_MENTION_COED_NONSACEM}

{START_BDO_COMMENTAIRE_NONSACEM}
Le présent dépôt porte sur [PartsSACEM]% de l'œuvre (parts de [ListeAC_SACEM] et leurs éditeurs [ListeE_SACEM]). Les autres [PartsNonSACEM]% reviennent à [ListeNonSACEM][EditeurDe], membre[Pluriel] d'une société de gestion collective étrangère.
{END_BDO_COMMENTAIRE_NONSACEM}
```

### CCEOM_template_univ.docx

- **Article 11** : Ajouter `{Mention_Article11_NonSACEM}` après `{tabcreasplit}`
- **Article 16** : Remplacer la mention en dur par `{Mention_Article16_NonSACEM}`

### COED_template_univ.docx

- **Article 3** : Ajouter `{Mention_COED_NonSACEM}` après `[editsplit]`

---

## 🔧 Logique de détection

Un ayant droit est considéré **NON-SACEM** si :
```
SocieteGestion != "SACEM" ET SocieteGestion n'est pas vide
```

Un ayant droit a une **part inédite** si :
```
Role = A, C, AC, AR ou AD
ET il est seul dans son lettrage (pas d'éditeur dans le même groupe)
```

---

## 📊 Résumé

| Catégorie | Nombre de balises |
|-----------|-------------------|
| Calcul | 4 |
| Listes SACEM | 4 |
| Listes NON-SACEM | 4 |
| Pluriel/Singulier | 4 |
| Spécifiques | 2 |
| Détection | 1 |
| **TOTAL** | **19** |

---

## 📅 Historique

- **25/01/2026** : Création initiale des 19 balises NON-SACEM
