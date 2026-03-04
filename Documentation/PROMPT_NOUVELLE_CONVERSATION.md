# PROJET SACEM GENERATOR VB.NET - Prompt de continuation

## 📂 Code source
**GitHub** : https://github.com/Exxon/SACEMGenerator

---

## 🎯 Description du projet
Application Windows Forms VB.NET pour générer automatiquement les documents contractuels SACEM à partir d'un fichier JSON :
- **BDO** (PDF) - Bulletin de Déclaration d'Œuvre (template officiel Bdo711.pdf, rempli via Python/pypdf)
- **CCEOM** (DOCX) - Contrat de Cession et d'Édition d'Œuvre Musicale
- **CCDAA** (DOCX) - Contrat de Cession du Droit d'Adaptation Audiovisuelle
- **COED** (DOCX) - Contrat de Co-Édition
- **Split Sheet** (DOCX) - Lettre de Répartition

---

## 🏗️ Architecture clé

| Fichier | Rôle |
|---------|------|
| `MainForm.vb` | Interface + statistiques (A/C/E, NON-SACEM, parts inédites) + alerte œuvre mixte |
| `Models/SACEMData.vb` | Classes de données (Identite, BDO, Adresse, Contact, AyantDroit) |
| `Services/SACEMJsonReader.vb` | Parsing JSON → SACEMData (inclut champ `SocieteGestion`) |
| `Services/BalisesGenerator.vb` | Génère 50+ balises simples dont 19 balises NON-SACEM |
| `Services/SuperbaliseGenerator.vb` | Génère superbalises `{xxx}` (paragraphes complexes) |
| `Services/ContractGenerator.vb` | Génération DOCX avec OpenXML + fusion runs fragmentés |
| `Services/BDOPdfGenerator.vb` | Génération PDF via script Python + auto-ajustement police |
| `Services/TableGenerator.vb` | Génération tableaux Word (signatures, répartition) |
| `Services/ParagraphTemplateReader.vb` | Lecture blocs `{START_xxx}...{END_xxx}` depuis template_paragrahs.docx |
| `Templates/template_paragrahs.docx` | Blocs de texte réutilisables avec formatage |

---

## 📋 Concepts métier SACEM essentiels

### Types de droits
| Type | Abréviation | Répartition | Modifiable ? |
|------|-------------|-------------|--------------|
| Exécution Publique | DEP | 4/12 - 4/12 - 4/12 (A-C-E) | ❌ Statutaire |
| Radio Mécanique | DR | 25% - 25% - 50% (A-C-E) | ❌ Statutaire (option 2019) |
| Reproduction Mécanique | DRM/Phono/PH | Libre (contractuel) | ✅ Négociable |

### Lettrage
Système de groupement : un éditeur doit avoir le **même lettrage** que le créateur qu'il édite.

### Part inédite
AC **seul dans son lettrage** (pas d'éditeur) = part inédite. Indépendant de SACEM/NON-SACEM.

### EACA (Éditeur à Compte d'Auteur)
Quand `Managelic` ou `Managesub` = nom du créateur → templates `SUBEAC`/`LICEAC` au lieu de `SUBS`/`LIC`.

### Œuvre mixte
Œuvre avec ayants droit SACEM + NON-SACEM (GEMA, KODA, PRS...).
- Détection : `SocieteGestion != "SACEM"` et non vide
- AC NON-SACEM → ne signe PAS CCEOM/CCDAA, signe Split Sheet
- E NON-SACEM → ne signe PAS COED, signe Split Sheet
- Mentions obligatoires Articles 11/16 CCEOM et Article 3 COED

---

## ⚠️ PIÈGES CRITIQUES À ÉVITER

### 1. Fragmentation des Runs Word
Word découpe le texte en plusieurs `<w:r>`. Une balise `[Faitle]` peut devenir `[Fait` + `le]`.
**Solution** : Appeler `MergeFragmentedRuns()` AVANT tout remplacement.

### 2. Champ SocieteGestion non lu
Le champ existait dans le modèle mais n'était pas lu dans `SACEMJsonReader.vb`.
**Solution** : Vérifier que TOUS les champs du modèle sont lus depuis le JSON.

### 3. Designation vide pour personnes physiques
Les personnes physiques n'ont pas de `Designation` (réservé aux morales).
**Solution** : Utiliser `Nom_Prenom` comme clé unique pour les physiques.

### 4. Count() LINQ en VB.NET
`List.Count(Function...)` ne fonctionne pas.
**Solution** : Utiliser `.Where(Function...).Count()`

### 5. Interpolation de chaîne VB.NET
`$"{vbCrLf}{'='}"` provoque des erreurs.
**Solution** : Utiliser concaténation classique `& vbCrLf &`

### 6. Références fichiers fantômes dans .vbproj
Si fichier supprimé manuellement, la référence reste dans `.vbproj`.
**Solution** : Ouvrir `.vbproj` et supprimer la ligne manuellement.

---

## ✅ État actuel (26/01/2026)

- [x] Balises simples et indexées fonctionnelles
- [x] Superbalises `{auteurspart}`, `{editeurspart}`, `{subpart}`, `{licpart}`
- [x] Tables `{tabsignature}`, `{tabcreasplit}`, `{tabcreasplit2}`
- [x] 19 balises NON-SACEM implémentées
- [x] Détection et affichage œuvre mixte dans MainForm
- [x] Alerte MessageBox listant les non-signataires
- [x] Affichage `(SOCIÉTÉ)` dans BDO pour NON-SACEM
- [x] Auto-ajustement police champ Commentaire BDO (5-10pt)
- [x] Statistiques UI : A/C/E, lettrages, NON-SACEM, parts inédites

---

## 🔧 Dépendances

- .NET Framework 4.7.2+
- DocumentFormat.OpenXml (NuGet)
- Newtonsoft.Json (NuGet)
- Python 3.x + pypdf (`pip install pypdf`)

---

## 📝 MA DEMANDE

**Avant de répondre, consulte le GitHub pour voir le code actuel.**

[DÉCRIS ICI TON PROBLÈME OU CE QUE TU VEUX FAIRE]

---

## 📚 Documentation complète

Une documentation technique détaillée (900+ lignes) a été créée couvrant :
- Contexte métier SACEM (acteurs, lettrage, répartition DEP/DR/DRM, EACA)
- Architecture technique complète
- Système de balises (simples, indexées, calculées, superbalises)
- Tous les pièges résolus avec solutions
- Fonctionnalité NON-SACEM complète

**Demande-la si besoin.**
