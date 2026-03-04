# 📚 SACEM GENERATOR VB.NET - Documentation Technique Complète

## 📋 Table des matières
1. [Présentation du projet](#1-présentation-du-projet)
2. [Contexte métier SACEM](#2-contexte-métier-sacem)
3. [Architecture et fichiers](#3-architecture-et-fichiers)
4. [Structure JSON des données](#4-structure-json-des-données)
5. [Système de balises](#5-système-de-balises)
6. [⚠️ PIÈGES ET PROBLÈMES RÉSOLUS](#6-️-pièges-et-problèmes-résolus)
7. [Fonctionnalité NON-SACEM](#7-fonctionnalité-non-sacem)
8. [Génération des documents](#8-génération-des-documents)
9. [Prompt de démarrage](#9-prompt-de-démarrage)

---

# 1. Présentation du projet

## 🎯 Objectif
Application Windows Forms VB.NET pour générer automatiquement les documents contractuels SACEM à partir d'un fichier JSON contenant les informations de l'œuvre et des ayants droit.

## 📄 Documents générés

| Abréviation | Nom complet | Description |
|-------------|-------------|-------------|
| **BDO** | Bulletin de Déclaration d'Œuvre | PDF officiel SACEM (Bdo711.pdf) |
| **CCEOM** | Contrat de Cession et d'Édition d'Œuvre Musicale | Contrat AC ↔ Éditeur |
| **CCDAA** | Contrat de Cession du Droit d'Adaptation Audiovisuelle | Droits synchro |
| **COED** | Contrat de Co-Édition | Contrat entre Éditeurs |
| **Split Sheet** | Lettre de Répartition | Liste tous les ayants droit (SACEM + étranger) |

## 🔧 Technologies
- **VB.NET** (.NET Framework 4.7.2+)
- **DocumentFormat.OpenXml** - Manipulation DOCX
- **Newtonsoft.Json** - Parsing JSON
- **Python + pypdf** - Remplissage PDF (appelé via Process)

---

# 2. Contexte métier SACEM

## 2.1 Qu'est-ce que la SACEM ?

La **SACEM** (Société des Auteurs, Compositeurs et Éditeurs de Musique) est une société de gestion collective française. Elle :
- Collecte les droits d'auteur auprès des diffuseurs (radios, TV, concerts, streaming...)
- Répartit ces droits aux ayants droit (auteurs, compositeurs, éditeurs)
- Gère un répertoire de millions d'œuvres musicales

## 2.2 Les acteurs d'une œuvre musicale

### 🎭 Les Créateurs (AC = Auteurs/Compositeurs)

| Rôle | Code | Description | Exemple |
|------|------|-------------|---------|
| **Auteur** | A | Écrit les paroles (texte) | Parolier |
| **Compositeur** | C | Compose la musique (mélodie) | Musicien |
| **Auteur-Compositeur** | AC | Fait les deux | Artiste solo |
| **Adaptateur** | AD | Adapte un texte existant | Traducteur de paroles |
| **Arrangeur** | AR | Arrange une composition existante | Orchestrateur |

### 🏢 Les Éditeurs (E)

| Rôle | Code | Description |
|------|------|-------------|
| **Éditeur** | E | Exploite commercialement l'œuvre |

**Fonction de l'éditeur :**
- Promotion et placement de l'œuvre (synchro pub/film, radio...)
- Administration des droits
- Avance financière aux créateurs
- En échange : perçoit une part des droits (généralement 50% de la part éditeur)

### 📊 Répartition des droits SACEM

La SACEM gère plusieurs types de droits avec des règles de répartition différentes. Il est **CRUCIAL** de bien distinguer ces types car ils n'obéissent pas aux mêmes règles.

---

#### A) DEP - Droits d'Exécution Publique (Statutaire - 4/12èmes)

Les DEP concernent : concerts live, radio, TV, lieux publics (bars, commerces...), streaming (part exécution).

**Règle STATUTAIRE (imposée par le Règlement Général SACEM) :**

```
╔═══════════════════════════════════════════════════════════════════════╗
║                    ŒUVRE ÉDITÉE (avec éditeur)                        ║
╠═════════════════╦═════════════════╦═══════════════════════════════════╣
║  4/12 AUTEURS   ║ 4/12 COMPOSITEURS ║      4/12 ÉDITEURS              ║
║    (33,33%)     ║     (33,33%)      ║        (33,33%)                 ║
╚═════════════════╩═════════════════╩═══════════════════════════════════╝

╔═══════════════════════════════════════════════════════════════════════╗
║                    ŒUVRE INÉDITE (sans éditeur)                       ║
╠═════════════════════════════════╦═════════════════════════════════════╣
║       6/12 AUTEURS              ║        6/12 COMPOSITEURS            ║
║          (50%)                  ║            (50%)                    ║
╚═════════════════════════════════╩═════════════════════════════════════╝
```

**⚠️ IMPORTANT** : 
- Cette répartition est **OBLIGATOIRE** et **NON NÉGOCIABLE** par contrat
- Fixée par les articles du Règlement Général de la SACEM
- Le partage **à l'intérieur** de chaque catégorie est **égalitaire par défaut**
- Depuis le 1er janvier 2019, option possible de suivre la clé DRM au sein de chaque catégorie

**Cas particuliers DEP :**
- **Œuvre instrumentale** (sans paroles) : le compositeur perçoit la part auteur
- **Œuvre a cappella** (sans musique) : l'auteur perçoit la part compositeur
- **Avec arrangeur** : 1/12 prélevé sur la part compositeur (2/12 pour musique symphonique de film)

---

#### B) DRM - Droits de Reproduction Mécanique / "Phono" (Contractuel = PH)

Les DRM (aussi appelés "droits phono") concernent : CD, vinyles, téléchargements, streaming (part reproduction).

**Règle CONTRACTUELLE (librement négociable) :**

```
╔═══════════════════════════════════════════════════════════════════════╗
║                    100% DES DRM                                       ║
║         Répartition LIBRE selon le contrat                            ║
║     Négociée entre les ayants droit et inscrite au BDO                ║
╚═══════════════════════════════════════════════════════════════════════╝
```

**PH = Pourcentage Horizontal** : C'est le pourcentage contractuel de chaque ayant droit pour les droits de reproduction mécanique. Il est inscrit dans la colonne "Clé Phono" du BDO.

**⚠️ IMPORTANT** : 
- Le PH est **LIBREMENT NÉGOCIABLE** entre les parties
- Le total doit faire 100%
- Il peut être égalitaire ou inégalitaire
- C'est CE pourcentage qu'on remplit dans le BDO

**Dans le JSON** : Le champ `BDO.PH` contient ce pourcentage pour chaque ayant droit.

---

#### C) DR - Droits Radio Mécaniques (Statutaire, option contractuelle depuis 2019)

Les droits "Radio Mécaniques" (DR) concernent : diffusion radio/TV qui implique une reproduction mécanique préalable (le support lu pour diffuser).

**Règle STATUTAIRE par défaut (articles 76 et suivants du Règlement Général) :**

```
╔═══════════════════════════════════════════════════════════════════════╗
║                    ŒUVRE ÉDITÉE - DR par défaut                       ║
╠═════════════════╦═════════════════╦═══════════════════════════════════╣
║  25% AUTEURS    ║ 25% COMPOSITEURS ║        50% ÉDITEURS              ║
║    (3/12)       ║     (3/12)       ║          (6/12)                  ║
╚═════════════════╩═════════════════╩═══════════════════════════════════╝

╔═══════════════════════════════════════════════════════════════════════╗
║                    ŒUVRE INÉDITE - DR par défaut                      ║
╠═════════════════════════════════╦═════════════════════════════════════╣
║         50% AUTEURS             ║         50% COMPOSITEURS            ║
╚═════════════════════════════════╩═════════════════════════════════════╝
```

**Option depuis le 1er janvier 2019** : Les ayants droit peuvent cocher une case sur le BDO (rubrique D) pour que les DR suivent la répartition DRM (clé Phono) **au sein de chaque catégorie**, sans déroger au partage entre catégories.

---

#### D) Partage Égalitaire vs Inégalitaire

##### Partage ÉGALITAIRE (par défaut SACEM)

```
À l'intérieur de chaque catégorie, tous les membres ont la même part.

Exemple DEP avec 2 auteurs, 1 compositeur, 2 éditeurs :
AUTEURS (4/12 = 33,33%) :
├── Auteur 1 : 33,33% ÷ 2 = 16,67%
└── Auteur 2 : 33,33% ÷ 2 = 16,67%

COMPOSITEURS (4/12 = 33,33%) :
└── Compositeur 1 : 33,33% (seul)

ÉDITEURS (4/12 = 33,33%) :
├── Éditeur 1 : 33,33% ÷ 2 = 16,67%
└── Éditeur 2 : 33,33% ÷ 2 = 16,67%
```

##### Partage INÉGALITAIRE (sur décision commune)

Le partage inégalitaire s'applique **au sein d'une même catégorie** pour les DRM et optionnellement pour les DEP/DR depuis 2019.

```
Exemple DRM avec 3 auteurs (inégalitaire négocié) :
├── Auteur 1 (auteur principal)     : 15,00%
├── Auteur 2 (co-auteur)            : 8,00%
└── Auteur 3 (contribution mineure) : 2,00%
                                      ────────
                       Total Auteurs : 25,00%
```

**⚠️ Conditions pour un partage inégalitaire :**
1. **Accord unanime** de tous les ayants droit concernés
2. **Case cochée** sur le BDO : "Partage inégalitaire"
3. **PH détaillé** pour chaque ayant droit dans le BDO

**Dans le JSON** : Le champ `Inegalitaire` indique si le partage est inégalitaire :
```json
{
  "Inegalitaire": "TRUE"  // ou "FALSE", "OUI", "NON", "X"
}
```

**Dans le BDO** : La case "Partage inégalitaire" est cochée automatiquement par le code si `Inegalitaire = TRUE`.

---

#### E) Tableau récapitulatif complet

| Type de droit | Abréviation | Répartition par défaut | Modifiable par contrat ? | Référence |
|---------------|-------------|------------------------|--------------------------|-----------|
| **Exécution Publique** | DEP | 4/12 - 4/12 - 4/12 | ❌ NON (statutaire) | Règlement Général |
| **Reproduction Mécanique** | DRM / Phono | Libre (PH) | ✅ OUI | BDO colonne "Clé Phono" |
| **Radio Mécanique** | DR | 25% - 25% - 50% | ✅ Option BDO depuis 2019 | Art. 76+ Règlement |
| **Intra-catégorie DEP** | - | Égalitaire | ✅ Option depuis 2019 | Case BDO rubrique D |
| **Intra-catégorie DRM** | - | Égalitaire | ✅ OUI (PH par ayant droit) | BDO |

---

#### F) Exemple complet réel

**ŒUVRE : "Ma Chanson" - 2 auteurs, 1 compositeur, 2 éditeurs - Partage DRM inégalitaire**

```
╔═══════════════════════════════════════════════════════════════════════╗
║                    DEP - DROITS D'EXÉCUTION PUBLIQUE                  ║
║                    (Statutaire 4/12 - 4/12 - 4/12)                    ║
╠═══════════════════╦═══════════════════╦═══════════════════════════════╣
║ AUTEURS 33,33%    ║ COMPOSITEURS 33,33%║ ÉDITEURS 33,33%              ║
╠═══════════════════╬═══════════════════╬═══════════════════════════════╣
║ DUPONT    16,67%  ║ DURAND    33,33%  ║ MUSIC ED  16,67%             ║
║ MARTIN    16,67%  ║                   ║ SONY      16,67%             ║
╠═══════════════════╩═══════════════════╩═══════════════════════════════╣
║ ⚠️ Partage égalitaire OBLIGATOIRE au sein de chaque catégorie        ║
║    (sauf option cochée rubrique D depuis 2019)                        ║
╚═══════════════════════════════════════════════════════════════════════╝

╔═══════════════════════════════════════════════════════════════════════╗
║                    DR - DROITS RADIO MÉCANIQUES                       ║
║                    (Statutaire 25% - 25% - 50%)                       ║
╠═══════════════════╦═══════════════════╦═══════════════════════════════╣
║ AUTEURS 25%       ║ COMPOSITEURS 25%  ║ ÉDITEURS 50%                 ║
╠═══════════════════╬═══════════════════╬═══════════════════════════════╣
║ DUPONT    12,50%  ║ DURAND    25,00%  ║ MUSIC ED  25,00%             ║
║ MARTIN    12,50%  ║                   ║ SONY      25,00%             ║
╚═══════════════════╩═══════════════════╩═══════════════════════════════╝

╔═══════════════════════════════════════════════════════════════════════╗
║                    DRM - DROITS REPRODUCTION MÉCANIQUE                ║
║                    (Contractuel = PH négocié librement)               ║
╠═══════════════════════════════════════════════════════════════════════╣
║ DUPONT Jean (A)       :  12,00%  ← négocié                           ║
║ MARTIN Marie (A)      :   8,00%  ← négocié                           ║
║ DURAND Pierre (C)     :  20,00%  ← négocié                           ║
║ MUSIC EDITION (E)     :  35,00%  ← négocié                           ║
║ SONY MUSIC (E)        :  25,00%  ← négocié                           ║
╠═══════════════════════════════════════════════════════════════════════╣
║                         TOTAL PH : 100,00%                            ║
║ ✅ Ces valeurs sont inscrites dans la colonne "Clé Phono" du BDO     ║
╚═══════════════════════════════════════════════════════════════════════╝
```

---

#### G) Ce qui est géré par le code (BalisesGenerator.vb)

| Donnée | Source JSON | Balise générée | Usage |
|--------|-------------|----------------|-------|
| Part de chaque ayant droit | `BDO.PH` | `[PH1]`, `[PH2]`... | Colonne BDO |
| Partage inégalitaire | `Inegalitaire` | Case cochée BDO | Champ 80 du PDF |
| Total parts SACEM | Calcul | `[PartsSACEM]` | Commentaires |
| Total parts NON-SACEM | Calcul | `[PartsNonSACEM]` | Commentaires |

**Note** : Le code ne calcule pas la répartition DEP/DR (c'est la SACEM qui le fait lors de la répartition). Le code ne manipule que le **PH contractuel** inscrit dans le JSON.

## 2.3 Le système de Lettrage

### Concept
Le **lettrage** est un système de groupement des ayants droit. Chaque groupe est identifié par une lettre (A, B, C, D...).

### Règle fondamentale
> **Un éditeur qui édite un créateur doit avoir le même lettrage que ce créateur.**

### Exemple concret

```
ŒUVRE : "Ma Chanson"

Lettrage A :
├── DUPONT Jean (Auteur) ──────────── A
└── MUSIC EDITION (Éditeur) ────────── A    ← édite DUPONT

Lettrage B :
├── MARTIN Marie (Compositeur) ────── B
└── MUSIC EDITION (Éditeur) ────────── B    ← édite aussi MARTIN
└── SONY MUSIC (Co-Éditeur) ─────────── B    ← co-édite avec MUSIC EDITION

Lettrage C :
└── DURAND Pierre (Auteur-Comp) ───── C    ← PAS D'ÉDITEUR = PART INÉDITE
```

### Part inédite
Quand un créateur n'a **pas d'éditeur dans son lettrage**, sa part reste "inédite" :
- Il garde 100% de sa part (pas de partage avec éditeur)
- Mention spéciale dans le BDO : "(part inédite)"

## 2.4 Calcul des parts (PH = Pourcentage Horizontal)

### Exemple de répartition

```
ŒUVRE avec 3 créateurs et 2 éditeurs :

CRÉATEURS :
- DUPONT Jean (A)     : 12,50%
- MARTIN Marie (C)    : 12,50%  
- DURAND Pierre (AC)  : 25,00%   ← Auteur ET Compositeur = double part
                        ────────
                        50,00%   (moitié créateurs)

ÉDITEURS :
- MUSIC EDITION (E)   : 25,00%
- SONY MUSIC (E)      : 25,00%
                        ────────
                        50,00%   (moitié éditeurs)

TOTAL                 : 100,00%
```

### Formule générale
```
Part Créateur = (100% ÷ 2) ÷ Nombre de créateurs
Part Éditeur = (100% ÷ 2) ÷ Nombre d'éditeurs

Ajustement si AC (Auteur-Compositeur) : compte pour 2 créateurs
```

## 2.5 Les contrats SACEM

### 📄 BDO - Bulletin de Déclaration d'Œuvre

**Objectif** : Déclarer officiellement l'œuvre à la SACEM

**Contenu** :
- Informations de l'œuvre (titre, durée, genre...)
- Liste complète des ayants droit avec leurs parts
- Signatures de tous les ayants droit SACEM

**Règle** : Seuls les membres SACEM signent le BDO. Les membres de sociétés étrangères sont mentionnés mais ne signent pas.

---

### 📄 CCEOM - Contrat de Cession et d'Édition d'Œuvre Musicale

**Objectif** : Formaliser la relation entre un créateur et son éditeur

**Parties** :
- L'AUTEUR et/ou COMPOSITEUR (cédant)
- L'ÉDITEUR (cessionnaire)

**Contenu clé** :
- **Article 1** : Identification de l'œuvre
- **Article 2** : Cession des droits d'édition
- **Article 3** : Durée de la cession
- **Article 8** : Rémunération de l'auteur
- **Article 11** : Mentions spéciales (co-édition, NON-SACEM...)
- **Article 16** : Dispositions particulières

**Règle NON-SACEM** : Si l'œuvre implique des membres NON-SACEM, une mention est ajoutée aux Articles 11 et 16 pour préciser que les pourcentages s'entendent au prorata de la part SACEM.

---

### 📄 CCDAA - Contrat de Cession du Droit d'Adaptation Audiovisuelle

**Objectif** : Autoriser l'utilisation de l'œuvre en synchronisation (pub, film, série, jeu vidéo...)

**Parties** : Mêmes que CCEOM

**Spécificité** : Concerne les droits "synchro" qui sont négociés séparément des droits d'exécution publique.

**Article 4** : Fait référence à l'Article 16 du CCEOM pour les mentions NON-SACEM.

---

### 📄 COED - Contrat de Co-Édition

**Objectif** : Organiser le partage entre plusieurs éditeurs

**Parties** :
- ÉDITEUR 1 (souvent l'éditeur original)
- ÉDITEUR 2, 3... (co-éditeurs)

**Contenu clé** :
- Répartition des parts entre éditeurs
- Qui administre l'œuvre (éditeur principal)
- Territoires de chaque éditeur

**Article 3** : Mention NON-SACEM si un des co-éditeurs est membre d'une société étrangère.

---

### 📄 Split Sheet - Lettre de Répartition

**Objectif** : Document international listant TOUS les ayants droit

**Particularité** : 
- Signé par TOUS (SACEM + NON-SACEM)
- Sert de preuve de la répartition convenue
- Utilisé pour les œuvres internationales

## 2.6 Sous-édition et Licences

### Sous-édition (Managesub)

**Concept** : Un éditeur français peut confier l'exploitation de l'œuvre à un éditeur étranger pour certains territoires.

**Exemple** :
```
MUSIC EDITION (France) ──sous-édite──► WARNER CHAPPELL (USA)
                                       pour le territoire américain
```

**Dans le JSON** : Le champ `Managesub` indique qui gère la sous-édition.

### Licences (Managelic)

**Concept** : L'éditeur peut accorder des licences d'exploitation (synchro, compilation...).

**Dans le JSON** : Le champ `Managelic` indique qui gère les licences.

### EACA - Éditeur à Compte d'Auteur

**Concept** : Situation où le créateur est aussi son propre éditeur, ou quand l'éditeur qui gère les droits est le créateur lui-même.

**Cas concrets :**

1. **Auto-édition** : L'auteur crée sa propre structure d'édition
   ```
   DUPONT Jean (Auteur) ──est aussi──► DUPONT MUSIC (Éditeur)
   ```

2. **Gestion par le créateur** : Un éditeur externe existe mais c'est le créateur qui gère
   ```
   Managelic = "DUPONT Jean"  (le créateur gère lui-même les licences)
   Managesub = "DUPONT Jean"  (le créateur gère lui-même la sous-édition)
   ```

**Détection dans le code :**
```vb
' EACA = Le gestionnaire (Managelic/Managesub) est le même que le créateur
Dim isEACA As Boolean = (managelic = creatorName) Or (managesub = creatorName)
```

**Particularité contractuelle :**
> Dans le cas EACA, c'est le créateur "qui se paye lui-même" pour la partie édition. Les paragraphes de sous-édition et licences utilisent des templates spécifiques (`SUBEAC`, `LICEAC`) avec une formulation adaptée.

**Impact sur les templates :**

| Situation | Template utilisé | Formulation |
|-----------|------------------|-------------|
| Éditeur externe | `SUBS` / `LIC` | "L'Éditeur X percevra..." |
| EACA | `SUBEAC` / `LICEAC` | "L'Auteur percevra directement..." |

**Exemple JSON EACA :**
```json
{
  "Identite": {
    "Designation": "",
    "Type": "Physique",
    "Nom": "DUPONT",
    "Prenom": "Jean"
  },
  "BDO": {
    "Role": "AC",
    "Lettrage": "A",
    "Managelic": "DUPONT Jean",    // ← Lui-même = EACA
    "Managesub": "DUPONT Jean"     // ← Lui-même = EACA
  }
}
```

## 2.7 Sociétés de gestion collective étrangères

### Principales sociétés

| Pays | Société | Type |
|------|---------|------|
| 🇫🇷 France | SACEM | Droits d'auteur |
| 🇩🇪 Allemagne | GEMA | Droits d'auteur |
| 🇬🇧 Royaume-Uni | PRS | Droits d'auteur |
| 🇺🇸 États-Unis | ASCAP, BMI, SESAC | Droits d'auteur |
| 🇩🇰 Danemark | KODA | Droits d'auteur |
| 🇪🇸 Espagne | SGAE | Droits d'auteur |
| 🇮🇹 Italie | SIAE | Droits d'auteur |
| 🇧🇪 Belgique | SABAM | Droits d'auteur |
| 🇳🇱 Pays-Bas | BUMA/STEMRA | Droits d'auteur |

### Accords de réciprocité

Les sociétés ont des accords entre elles :
- La SACEM collecte en France pour les membres GEMA, PRS, etc.
- La GEMA collecte en Allemagne pour les membres SACEM, etc.

**Conséquence** : Un membre GEMA n'a pas besoin de signer les contrats SACEM français, ses droits sont protégés via l'accord SACEM-GEMA.

## 2.8 Œuvre mixte

### Définition officielle SACEM
> Une **œuvre mixte** est une œuvre dont les ayants droit sont membres de sociétés de gestion collective différentes.

### Exemple
```
ŒUVRE : "International Hit"

SACEM (France) :
├── DUPONT Jean (Auteur)        : 25%
└── MUSIC EDITION (Éditeur)     : 25%
                                  ────
                                  50% SACEM

GEMA (Allemagne) :
├── MÜLLER Hans (Compositeur)   : 25%
└── BERLIN MUSIK (Éditeur)      : 25%
                                  ────
                                  50% NON-SACEM (du point de vue SACEM)

→ C'est une ŒUVRE MIXTE
```

### Règles pour les œuvres mixtes

1. **BDO** : Mentionne tous les ayants droit, mais seuls les SACEM signent
2. **Commentaire BDO** : Précise la répartition SACEM / NON-SACEM
3. **CCEOM/CCDAA** : Les pourcentages sont "au prorata de la part SACEM"
4. **COED** : Même logique si un co-éditeur est étranger
5. **Split Sheet** : Tout le monde signe

## 2.9 Règles de génération des documents

### Qui signe quoi ?

```
╔══════════════════╦═══════╦═════════╦═════════╦════════╦═════════════╗
║ Type d'ayant     ║  BDO  ║  CCEOM  ║  CCDAA  ║  COED  ║ Split Sheet ║
╠══════════════════╬═══════╬═════════╬═════════╬════════╬═════════════╣
║ AC SACEM         ║ ✅    ║ ✅      ║ ✅      ║   -    ║ ✅          ║
║ AC NON-SACEM     ║ mention║ ❌      ║ ❌      ║   -    ║ ✅          ║
║ E SACEM          ║ ✅    ║ ✅      ║ ✅      ║ ✅     ║ ✅          ║
║ E NON-SACEM      ║ mention║   -     ║   -     ║ ❌     ║ ✅          ║
╚══════════════════╩═══════╩═════════╩═════════╩════════╩═════════════╝
```

### Cas EACA (Éditeur à Compte d'Auteur)

Quand un créateur gère lui-même ses droits (auto-édition), les superbalises `{subpart}` et `{licpart}` utilisent des templates différents :

```
╔═══════════════════╦════════════════════╦═══════════════════════════════╗
║ Situation         ║ Template           ║ Formulation                   ║
╠═══════════════════╬════════════════════╬═══════════════════════════════╣
║ Éditeur externe   ║ START_SUBS         ║ "lui verse" / "leur verse"    ║
║ EACA              ║ START_SUBEAC       ║ "se verse" (à lui-même)       ║
╠═══════════════════╬════════════════════╬═══════════════════════════════╣
║ Éditeur externe   ║ START_LIC          ║ "lui verse" / "leur verse"    ║
║ EACA              ║ START_LICEAC       ║ "se verse" (à lui-même)       ║
╚═══════════════════╩════════════════════╩═══════════════════════════════╝
```

**Logique de détection EACA dans `SuperbaliseGenerator.vb` :**
```vb
' Pour chaque groupe de sous-édition/licence
Dim managerName As String = ayant.BDO.Managelic  ' ou Managesub
Dim creatorName As String = GetDisplayIdentifiant(ayant)

' Si le gestionnaire est le créateur lui-même = EACA
Dim isEACA As Boolean = managerName.ToUpper().Contains(creatorName.ToUpper())

If isEACA Then
    templateKey = "SUBEAC"  ' ou "LICEAC"
Else
    templateKey = "SUBS"    ' ou "LIC"
End If
```

### Mentions obligatoires (œuvre mixte)

**Dans CCEOM Article 11 :**
> "Il est ici précisé que l'ŒUVRE est coéditée avec d'une part, [ÉDITEURS SACEM], et d'autre part [ÉDITEURS NON-SACEM] ([SOCIÉTÉ]), qui en contrôlent respectivement [%SACEM]% et [%NON-SACEM]%."

**Dans CCEOM Article 16 :**
> "[NOM NON-SACEM] est membre d'une société de gestion collective autre que la SACEM. Il n'est pas signataire du présent contrat mais de la lettre de répartition ci-après annexée."

**Dans COED Article 3 :**
> Même logique que CCEOM pour les éditeurs étrangers.

## 2.10 Glossaire métier

| Terme | Définition |
|-------|------------|
| **Ayant droit** | Personne ou entité détenant des droits sur l'œuvre |
| **Répertoire** | Base de données des œuvres déclarées à la SACEM |
| **DEP** | Droits d'Exécution Publique (concerts, radio, TV...) |
| **DRM** | Droits de Reproduction Mécanique (CD, streaming...) |
| **Synchro** | Synchronisation (musique + image : pub, film...) |
| **Cession** | Transfert de droits du créateur vers l'éditeur |
| **Mandat** | Autorisation de gérer les droits (sans transfert de propriété) |
| **COAD** | Code Ayant Droit (identifiant SACEM) |
| **IPI** | Interested Parties Information (identifiant international) |
| **ISWC** | International Standard Musical Work Code (identifiant œuvre) |

---

# 3. Architecture et fichiers

```
SACEMGenerator/
├── MainForm.vb                    # Interface graphique principale
├── Models/
│   └── SACEMData.vb               # Modèle de données (classes)
├── Services/
│   ├── SACEMJsonReader.vb         # Lecture et parsing JSON
│   ├── BalisesGenerator.vb        # Génération des 50+ balises simples
│   ├── SuperbaliseGenerator.vb    # Génération des superbalises complexes
│   ├── ContractGenerator.vb       # Génération des contrats DOCX
│   ├── BDOPdfGenerator.vb         # Génération du BDO PDF
│   ├── TableGenerator.vb          # Génération des tableaux Word
│   └── ParagraphTemplateReader.vb # Lecture des blocs template_paragrahs.docx
├── Scripts/
│   └── FillBDO/
│       └── FillBDOFromJson.py     # Script Python pour remplir le PDF
├── Templates/
│   ├── Bdo711.pdf                 # Template PDF BDO officiel SACEM
│   ├── CCEOM_template_univ.docx   # Template contrat CCEOM
│   ├── CCDAA_template.docx        # Template contrat CCDAA
│   ├── COED_template_univ.docx    # Template contrat COED
│   └── template_paragrahs.docx    # Blocs de paragraphes réutilisables
└── Documentation/
    └── (ce fichier)
```

---

# 4. Structure JSON des données

## 3.1 Schéma global

```json
{
  "Titre": "NOM DE L'ŒUVRE",
  "SousTitre": "",
  "Interprete": "Nom de l'interprète",
  "Duree": "00:03:45",
  "Genre": "Variétés",
  "Date": "01/01/2024",
  "ISWC": "",
  "Lieu": "Paris",
  "Territoire": "Monde",
  "Arrangement": "Toutes",
  "Inegalitaire": "FALSE",
  "Commentaire": "",
  "Faita": "Paris",
  "Faitle": "15/01/2024",
  "Declaration": "EDITEUR X",
  "Format": "EDITEUR X",
  "AyantsDroit": [...]
}
```

## 3.2 Structure d'un Ayant Droit

```json
{
  "Identite": {
    "Designation": "NOM SOCIÉTÉ",       // Pour personnes morales
    "Type": "Physique" | "Moral",
    "Nom": "DUPONT",                     // Pour personnes physiques
    "Prenom": "Jean",
    "Pseudonyme": "DJ JEAN",
    "Genre": "MR" | "MME",               // Civilité pour physiques
    "Nele": "01/01/1980",                // Date naissance
    "Nea": "Paris",                      // Lieu naissance
    "SocieteGestion": "SACEM" | "GEMA" | "KODA" | etc.,
    "FormeJuridique": "SAS" | "SARL" | "EURL" | etc.,
    "Capital": "10000",
    "RCS": "Paris",
    "Siren": "123456789",
    "GenreRepresentant": "MR",
    "PrenomRepresentant": "Pierre",
    "NomRepresentant": "MARTIN",
    "FonctionRepresentant": "Gérant"
  },
  "BDO": {
    "Role": "A" | "C" | "AC" | "E" | "AR" | "AD",
    "COAD/IPI": "123456789",
    "PH": "8.33",                        // Pourcentage
    "Lettrage": "A",                     // Groupe de l'ayant droit
    "Managelic": "EDITEUR X",            // Qui gère les licences
    "Managesub": "EDITEUR X"             // Qui gère la sous-édition
  },
  "Adresse": {
    "NumVoie": "123",
    "TypeVoie": "rue",
    "NomVoie": "de la Musique",
    "CP": "75001",
    "Ville": "Paris",
    "Pays": "France"
  },
  "Contact": {
    "Mail": "contact@exemple.fr",
    "Tel": "0123456789"
  }
}
```

## 3.3 Rôles des ayants droit

| Code | Signification | Catégorie |
|------|---------------|-----------|
| **A** | Auteur (parolier) | Créateur (AC) |
| **C** | Compositeur | Créateur (AC) |
| **AC** | Auteur-Compositeur | Créateur (AC) |
| **AR** | Arrangeur | Créateur (AC) |
| **AD** | Adaptateur | Créateur (AC) |
| **E** | Éditeur | Éditeur (E) |

---

# 5. Système de balises

## 4.1 Types de balises

### A) Balises simples `[NomBalise]`
Remplacement direct par une valeur du JSON.

```
[Titre] → "Ma Chanson"
[Interprete] → "Artiste X"
[Faitle] → "15/01/2024"
```

### B) Balises indexées `[NomBalise1]`, `[NomBalise2]`...
Pour les ayants droit (max 17 lignes dans le BDO).

```
[Role1] → "A"
[Designation1] → "DUPONT Jean"
[Lettrage1] → "A"
[PH1] → "8,33"
```

### C) Balises calculées
Générées par logique à partir des données.

```
[auteurslist] → "Jean DUPONT, Marie MARTIN"
[compositeurslist] → "Pierre DURAND"
[editeurslist] → "EDITEUR X, EDITEUR Y"
[editsplit] → Répartition des parts éditeurs
```

### D) Superbalises `{NomSuperbalise}`
Blocs de texte complets avec formatage, générés dynamiquement.

```
{auteurspart} → Paragraphes complets pour chaque auteur
{editeurspart} → Paragraphes complets pour chaque éditeur
{tabsignature} → Tableau des signatures
{tabcreasplit} → Tableau de répartition créateurs
{subpart} → Paragraphes sous-édition
{licpart} → Paragraphes licences
```

## 4.2 Fichier template_paragrahs.docx

Ce fichier contient les **blocs de texte réutilisables** avec formatage (gras, italique).

### Structure des blocs :
```
{START_Physique}
En qualité [RoleGenre] [Civilite] [Nom] [Prenom]...
{END_Physique}

{START_EURL}
[Designation] [FormeJuridique], au capital social de...
{END_EURL}

{START_MENTION_ART11_NONSACEM}
Il est ici précisé que l'ŒUVRE est coéditée...
{END_MENTION_ART11_NONSACEM}
```

### Blocs disponibles :
| Bloc | Usage |
|------|-------|
| `Physique` | Paragraphe personne physique |
| `EURL` | Paragraphe EURL/SARL |
| `EI` | Paragraphe Entrepreneur Individuel |
| `ASS` | Paragraphe Association |
| `SA` | Paragraphe SA/SAS/SASU |
| `SUBS` | Paragraphe sous-édition |
| `SUBEAC` | Paragraphe sous-édition EACA |
| `LIC` | Paragraphe licences |
| `LICEAC` | Paragraphe licences EACA |
| `MENTION_ART11_NONSACEM` | Mention NON-SACEM Article 11 |
| `MENTION_ART16_NONSACEM` | Mention NON-SACEM Article 16 |
| `MENTION_COED_NONSACEM` | Mention NON-SACEM COED |
| `BDO_COMMENTAIRE_NONSACEM` | Commentaire BDO œuvre mixte |

---

# 6. ⚠️ PIÈGES ET PROBLÈMES RÉSOLUS

## 🔴 PIÈGE #1 : Fragmentation des Runs dans Word

### Le problème
Word découpe le texte en plusieurs `<w:r>` (runs) de manière imprévisible.
Une balise `[Faitle]` peut devenir :
```xml
<w:r><w:t>[Fait</w:t></w:r>
<w:r><w:t>le]</w:t></w:r>
```

### Symptôme
La balise n'est pas remplacée car le code cherche `[Faitle]` dans un seul run.

### Solution
**Fusion des runs avant remplacement** dans `ContractGenerator.vb` :

```vb
Private Sub MergeFragmentedRuns(body As Body)
    ' Parcourir tous les paragraphes
    For Each para In body.Descendants(Of Paragraph)()
        ' Collecter le texte de tous les runs
        Dim fullText As String = ""
        For Each run In para.Descendants(Of Run)()
            fullText &= run.InnerText
        Next
        
        ' Si le texte contient une balise fragmentée, reconstruire
        If fullText.Contains("[") AndAlso fullText.Contains("]") Then
            ' Supprimer les runs existants et créer un seul run
            ' avec le texte complet
        End If
    Next
End Sub
```

### Prévention
- Toujours appeler `MergeFragmentedRuns()` AVANT les remplacements
- Éviter de modifier manuellement les templates Word (copier-coller peut fragmenter)

---

## 🔴 PIÈGE #2 : Champ SocieteGestion non lu

### Le problème
Le champ `SocieteGestion` existait dans le modèle `SACEMData.vb` mais n'était **jamais lu** depuis le JSON dans `SACEMJsonReader.vb`.

### Symptôme
Tous les ayants droit apparaissaient comme SACEM même s'ils étaient GEMA/KODA.

### Solution
Ajouter dans `SACEMJsonReader.vb`, méthode `ParseAyantDroit()` :

```vb
' Société de gestion collective (SACEM, GEMA, KODA, etc.)
ayant.Identite.SocieteGestion = GetStringValue(identiteObj, "SocieteGestion")
```

### Prévention
Toujours vérifier que TOUS les champs du modèle sont effectivement lus depuis le JSON.

---

## 🔴 PIÈGE #3 : Designation vide pour personnes physiques

### Le problème
Les personnes physiques n'ont pas de `Designation` dans le JSON (c'est pour les personnes morales). Le code utilisait `Designation` comme clé unique, causant des collisions.

### Symptôme
Plusieurs auteurs fusionnés en un seul, ou noms vides.

### Solution
Créer une clé unique basée sur le type :

```vb
Private Function GetUniqueKey(ayant As AyantDroit) As String
    If ayant.Identite.Type = "Moral" Then
        Return ayant.Identite.Designation.ToUpper()
    Else
        Return $"{ayant.Identite.Nom}_{ayant.Identite.Prenom}".ToUpper()
    End If
End Function
```

---

## 🔴 PIÈGE #4 : Formatage perdu lors du remplacement

### Le problème
Le remplacement de texte supprimait le formatage (gras, italique, couleur) du template.

### Symptôme
Tout le texte en police normale après génération.

### Solution
**Hériter le formatage du run original** :

```vb
' Copier les propriétés du run original
Dim originalRunProps = originalRun.RunProperties?.CloneNode(True)
newRun.RunProperties = originalRunProps
```

---

## 🔴 PIÈGE #5 : Count() avec prédicat LINQ en VB.NET

### Le problème
`List.Count(Function(x) ...)` ne fonctionne pas directement en VB.NET.

### Symptôme
Erreur de compilation : "n'a aucun paramètre et son type de retour ne peut pas être indexé"

### Solution
Utiliser `.Where().Count()` :

```vb
' ❌ Ne fonctionne pas
countAC = roles.Count(Function(r) r = "A" OrElse r = "C")

' ✅ Fonctionne
countAC = roles.Where(Function(r) r = "A" OrElse r = "C").Count()
```

---

## 🔴 PIÈGE #6 : Interpolation de chaîne avec caractères spéciaux

### Le problème
En VB.NET, `$"{...}"` ne supporte pas certaines syntaxes comme `{'='}`.

### Symptôme
Erreur : "Expression attendue" ou "'}' attendu"

### Solution
Utiliser la concaténation classique :

```vb
' ❌ Ne fonctionne pas
txtLog.AppendText($"{'='}{New String("="c, 50)}")

' ✅ Fonctionne
txtLog.AppendText(New String("="c, 50) & vbCrLf)
```

---

## 🔴 PIÈGE #7 : End Sub manquant après copier-coller

### Le problème
Lors d'ajout de code par copier-coller, le `End Sub` peut être oublié.

### Symptôme
Erreur : "End Sub attendu" ou "Cette instruction ne peut pas apparaître dans le corps de méthode"

### Solution
Toujours vérifier la structure complète de la méthode après modification.

---

## 🔴 PIÈGE #8 : Référence de fichier dans .vbproj

### Le problème
Si on supprime un fichier manuellement (explorateur Windows), la référence reste dans le fichier `.vbproj`.

### Symptôme
Erreur de compilation : "Impossible de copier le fichier car il est introuvable"

### Solution
1. Ouvrir `.vbproj` avec un éditeur de texte
2. Chercher et supprimer la ligne référençant le fichier manquant
3. Ou dans Visual Studio : clic droit sur le fichier (icône ⚠️) → Exclure du projet

---

# 7. Fonctionnalité NON-SACEM

## 6.1 Concept d'œuvre mixte

Une **œuvre mixte** contient des ayants droit membres de sociétés de gestion différentes (ex: SACEM + GEMA + KODA).

## 6.2 Détection automatique

```vb
' Un ayant droit est NON-SACEM si :
Dim isSACEM As Boolean = (societe = "SACEM" OrElse String.IsNullOrEmpty(societe))
Dim isNonSACEM As Boolean = Not isSACEM
```

## 6.3 Part inédite

**Définition** : Un AC (Auteur/Compositeur) est en "part inédite" s'il est **seul dans son lettrage** (pas d'éditeur dans son groupe).

```vb
' Détection part inédite
Dim hasAC As Boolean = roles.Any(Function(r) r = "A" Or r = "C" Or r = "AC")
Dim hasE As Boolean = roles.Any(Function(r) r = "E")
Dim isPartInedite As Boolean = hasAC AndAlso Not hasE
```

**Important** : Part inédite est INDÉPENDANT de SACEM/NON-SACEM !

## 6.4 Les 19 balises NON-SACEM

### Calcul (4)
| Balise | Description | Exemple |
|--------|-------------|---------|
| `[PartsSACEM]` | % total SACEM | `91,68` |
| `[PartsNonSACEM]` | % total NON-SACEM | `8,32` |
| `[PartsSACEM_Texte]` | En lettres | `quatre-vingt-onze...` |
| `[PartsNonSACEM_Texte]` | En lettres | `huit virgule...` |

### Listes SACEM (4)
| Balise | Description |
|--------|-------------|
| `[ListeAC_SACEM]` | Auteurs/Compositeurs SACEM |
| `[ListeE_SACEM]` | Éditeurs SACEM |
| `[ListeSACEM]` | Tous SACEM (AC puis E) |
| `[ListeSACEM_Noms]` | Noms seuls SACEM |

### Listes NON-SACEM (4)
| Balise | Description | Format |
|--------|-------------|--------|
| `[ListeAC_NonSACEM]` | AC NON-SACEM | `NOM (SOCIÉTÉ)` |
| `[ListeE_NonSACEM]` | E NON-SACEM | `NOM (SOCIÉTÉ)` |
| `[ListeNonSACEM]` | Tous NON-SACEM | `NOM (SOCIÉTÉ)` |
| `[ListeNonSACEM_Noms]` | Noms seuls | `NOM` |

### Pluriel/Singulier (4)
| Balise | Singulier | Pluriel |
|--------|-----------|---------|
| `[EstSont]` | `est` | `sont` |
| `[IlIls]` | `Il` | `Ils` |
| `[NestPasNeSontPas]` | `n'est pas` | `ne sont pas` |
| `[Pluriel]` | `` | `s` |

### Spécifiques (3)
| Balise | Description |
|--------|-------------|
| `[NumPoint]` | Numéro du point Article 16 |
| `[EditeurDe]` | " éditeur de NOM" si applicable |
| `[HasNonSACEM]` | `True` ou `False` |

## 6.5 Affichage dans le BDO

### Tableau des ayants droit
```
SACEM :           "DUPONT Jean"
NON-SACEM :       "BJERG Oliver (KODA)"
Part inédite :    "MARTIN Pierre (part inédite)"
NON-SACEM + PI :  "BJERG Oliver (KODA – part inédite)"
```

### Commentaire BDO (œuvre mixte)
> "Le présent dépôt porte sur 91,68% de l'œuvre (parts de DUPONT Jean, MARTIN Pierre et leurs éditeurs EDITEUR X, EDITEUR Y). Les autres 8,32% reviennent à BJERG Oliver (KODA), membre d'une société de gestion collective étrangère."

## 6.6 Documents et signatures

| Type d'ayant droit | BDO | CCEOM | CCDAA | COED | Split Sheet |
|--------------------|-----|-------|-------|------|-------------|
| AC SACEM | ✅ | ✅ signe | ✅ signe | - | ✅ signe |
| AC NON-SACEM | ✅ mention | ❌ ne signe pas | ❌ ne signe pas | - | ✅ signe |
| E SACEM | ✅ | ✅ signe | ✅ signe | ✅ signe | ✅ signe |
| E NON-SACEM | ✅ mention | - | - | ❌ ne signe pas | ✅ signe |

## 6.7 Alerte MessageBox

Au chargement d'un JSON avec membres NON-SACEM, une alerte s'affiche :

```
⚠ ŒUVRE MIXTE DÉTECTÉE

Les ayants droit suivants ne sont pas membres de la SACEM :

AUTEURS/COMPOSITEURS :
• Oliver Zoega BJERG (KODA)

→ Ne signeront PAS : CCEOM, CCDAA
→ Signeront UNIQUEMENT : Split Sheet

ÉDITEURS :
• CUBEATZ MUSIC GMBH (GEMA)

→ Ne signeront PAS : COED
→ Signeront UNIQUEMENT : Split Sheet
```

---

# 8. Génération des documents

## 7.1 Flux de génération

```
1. Charger JSON → SACEMJsonReader.LoadFromFile()
2. Générer balises simples → BalisesGenerator.GenerateAllBalises()
3. Générer superbalises → SuperbaliseGenerator
4. Générer tables → TableGenerator
5. Pour chaque contrat :
   a. Copier template → destination
   b. Fusionner runs fragmentés
   c. Remplacer balises [xxx]
   d. Remplacer superbalises {xxx}
   e. Insérer tables
   f. Sauvegarder
6. Pour BDO PDF :
   a. Générer JSON des valeurs
   b. Appeler script Python
   c. Remplir champs AcroForm
```

## 7.2 Auto-ajustement police BDO

Le champ Commentaire du BDO a une taille fixe. Le script Python calcule automatiquement la taille de police :

| Longueur texte | Taille police |
|----------------|---------------|
| Court | 10pt |
| Moyen | 8-9pt |
| Long | 6-7pt |
| Très long | 5pt (minimum) |

---

# 9. Prompt de démarrage

Utilisez ce prompt au début de chaque nouvelle conversation :

```
# PROJET SACEM GENERATOR VB.NET

## GitHub (code source)
https://github.com/Exxon/SACEMGenerator

## Description
Application Windows Forms VB.NET pour générer des documents SACEM :
- BDO (PDF) - Bulletin de Déclaration d'Œuvre via Python/pypdf
- CCEOM (DOCX) - Contrat de Cession et d'Édition
- CCDAA (DOCX) - Contrat Adaptation Audiovisuelle  
- COED (DOCX) - Contrat de Co-Édition
- Split Sheet (DOCX) - Lettre de Répartition

## Architecture clé
- MainForm.vb : Interface + statistiques (A/C/E, NON-SACEM, parts inédites)
- Services/BalisesGenerator.vb : 50+ balises dont 19 NON-SACEM
- Services/BDOPdfGenerator.vb : Génération PDF avec auto-ajustement police
- Services/ContractGenerator.vb : Génération DOCX avec OpenXML
- Services/SACEMJsonReader.vb : Parsing JSON (inclut SocieteGestion)
- Templates/template_paragrahs.docx : Blocs {START_xxx}...{END_xxx}

## ⚠️ Pièges connus
1. Runs fragmentés Word → MergeFragmentedRuns() obligatoire
2. Champ SocieteGestion → Vérifier qu'il est LU dans SACEMJsonReader
3. Designation vide pour Physiques → Utiliser Nom_Prenom comme clé
4. Count() LINQ → Utiliser .Where().Count() en VB.NET
5. Fichiers supprimés → Nettoyer le .vbproj

## Fonctionnalité NON-SACEM
- Détection : SocieteGestion != "SACEM"
- Part inédite : AC seul dans son lettrage (pas d'E)
- Affichage BDO : "NOM (SOCIÉTÉ)" pour NON-SACEM
- Alerte MessageBox listant les non-signataires
- Mentions automatiques dans CCEOM/COED

## État actuel (26/01/2026)
✅ Balises NON-SACEM (19) implémentées
✅ Statistiques MainForm complètes
✅ Alerte non-signataires
✅ Affichage (SOCIÉTÉ) dans BDO
✅ Auto-ajustement police commentaire BDO

## MA DEMANDE :
[Décrivez votre problème ou ce que vous voulez faire]

Consultez le GitHub pour voir le code actuel avant de répondre.
```

---

# 📅 Historique des versions

| Date | Version | Modifications |
|------|---------|---------------|
| 21/01/2026 | 1.0 | Version initiale |
| 22/01/2026 | 1.1 | Ajout BDO PDF, correction formatage |
| 24/01/2026 | 1.2 | Correction Designation physique |
| 25/01/2026 | 1.3 | Recherche NON-SACEM, spécification |
| 26/01/2026 | 1.4 | Implémentation NON-SACEM complète |

---

*Documentation générée le 26/01/2026*
*Projet : SACEM Generator VB.NET*
*GitHub : https://github.com/Exxon/SACEMGenerator*
