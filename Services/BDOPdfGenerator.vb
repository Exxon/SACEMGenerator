Imports System.IO
Imports System.Text

''' <summary>
''' Générateur de BDO SACEM par remplissage du PDF officiel (Bdo711.pdf)
''' Utilise iTextSharp ou un script Python pour remplir les champs AcroForm
''' </summary>
Public Class BDOPdfGenerator
    Private ReadOnly _data As SACEMData
    Private _log As New List(Of String)

    ' ============================================================
    ' MAPPING DES CHAMPS PDF BDO 711
    ' ============================================================

    ' Champs en-tête (Section A)
    Private Shared ReadOnly FIELD_MAPPING As New Dictionary(Of String, String) From {
        {"Titre", "1"},
        {"SousTitre", "2"},
        {"hh", "6"},
        {"mm", "7"},
        {"ss", "8"},
        {"Genre", "3"},
        {"ArrangementToutes", "Check Box77"},
        {"ArrangementDeterminees", "Check Box88"},
        {"Arrangement", "4"},
        {"JJ", "14c"},
        {"MM", "13c"},
        {"AAAA", "15c"},
        {"ISWC", "z20"},
        {"Lieu", "5"},
        {"TerritoireMonde", "Check Box7B"},
        {"TerritoireAutres", "Check Box8B"},
        {"Territoire", "Text1111"},
        {"Interprete", "9"},
        {"PartageInegalitaire", "Check Box80"},
        {"FJJ", "151"},
        {"FMM", "152"},
        {"FAAAA", "153"},
        {"Commentaire", "10"}
    }

    ''' <summary>
    ''' Log des opérations
    ''' </summary>
    Public ReadOnly Property GenerationLog As List(Of String)
        Get
            Return _log
        End Get
    End Property

    ''' <summary>
    ''' Constructeur
    ''' </summary>
    Public Sub New(data As SACEMData)
        _data = data
    End Sub

    ''' <summary>
    ''' Génère le BDO en remplissant le PDF officiel
    ''' </summary>
    ''' <param name="bdoPdfTemplatePath">Chemin vers Bdo711.pdf</param>
    ''' <param name="outputPath">Chemin du PDF de sortie</param>
    Public Function Generate(bdoPdfTemplatePath As String, outputPath As String) As Boolean
        Try
            _log.Clear()
            _log.Add("=== GÉNÉRATION BDO PDF SACEM ===")
            _log.Add($"Titre: {_data.Titre}")
            _log.Add($"Ayants droit: {_data.AyantsDroit.Count}")

            If Not File.Exists(bdoPdfTemplatePath) Then
                Throw New FileNotFoundException($"Template PDF BDO introuvable: {bdoPdfTemplatePath}")
            End If

            ' Générer les valeurs des champs
            Dim fieldValues As Dictionary(Of String, String) = GenerateFieldValues()
            _log.Add($"Champs préparés: {fieldValues.Count}")

            ' Créer le fichier JSON temporaire avec les valeurs
            Dim tempJsonPath As String = Path.Combine(Path.GetTempPath(), $"bdo_fields_{Guid.NewGuid()}.json")
            WriteFieldValuesJson(fieldValues, tempJsonPath)

            ' Appeler le script Python pour remplir le PDF
            Dim success As Boolean = FillPdfWithPython(bdoPdfTemplatePath, tempJsonPath, outputPath)

            ' Nettoyer
            Try
                If File.Exists(tempJsonPath) Then File.Delete(tempJsonPath)
            Catch
            End Try

            If success Then
                _log.Add("✓ BDO PDF généré avec succès")
                _log.Add($"Fichier: {outputPath}")
            End If

            Return success

        Catch ex As Exception
            _log.Add($"✗ ERREUR: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Génère toutes les valeurs des champs pour le PDF
    ''' </summary>
    Private Function GenerateFieldValues() As Dictionary(Of String, String)
        Dim values As New Dictionary(Of String, String)

        ' --- Section A : Informations de l'œuvre ---
        values(FIELD_MAPPING("Titre")) = If(_data.Titre, "")
        values(FIELD_MAPPING("SousTitre")) = If(_data.SousTitre, "")
        values(FIELD_MAPPING("Genre")) = If(_data.Genre, "")
        values(FIELD_MAPPING("ISWC")) = If(_data.ISWC, "")
        values(FIELD_MAPPING("Lieu")) = If(_data.Lieu, "")
        values(FIELD_MAPPING("Interprete")) = If(_data.Interprete, "")
        values(FIELD_MAPPING("Arrangement")) = If(_data.Arrangement, "")
        values(FIELD_MAPPING("Territoire")) = If(_data.Territoire, "")

        ' Durée : extraire hh:mm:ss
        If Not String.IsNullOrEmpty(_data.Duree) Then
            Dim parts = _data.Duree.Split(":"c)
            If parts.Length >= 3 Then
                values(FIELD_MAPPING("hh")) = parts(0)
                values(FIELD_MAPPING("mm")) = parts(1)
                values(FIELD_MAPPING("ss")) = parts(2)
            ElseIf parts.Length = 2 Then
                values(FIELD_MAPPING("mm")) = parts(0)
                values(FIELD_MAPPING("ss")) = parts(1)
            End If
        End If

        ' Date première exploitation
        Dim dateStr As String = If(_data.Date, "").Replace("""", "")
        If Not String.IsNullOrEmpty(dateStr) Then
            Dim parts = dateStr.Split("/"c)
            If parts.Length = 3 Then
                values(FIELD_MAPPING("JJ")) = parts(0)
                values(FIELD_MAPPING("MM")) = parts(1)
                values(FIELD_MAPPING("AAAA")) = parts(2)
            End If
        End If

        ' Cases à cocher - Territoire
        Dim territoire As String = If(_data.Territoire, "").ToLower()
        If territoire.Contains("monde") OrElse String.IsNullOrEmpty(territoire) Then
            values(FIELD_MAPPING("TerritoireMonde")) = "/Oui"
        End If
        If Not String.IsNullOrEmpty(territoire) AndAlso Not territoire.Contains("monde") Then
            values(FIELD_MAPPING("TerritoireAutres")) = "/Oui"
        End If

        ' Partage inégalitaire
        Dim inegalitaire As String = If(_data.Inegalitaire, "").ToUpper()
        If inegalitaire = "TRUE" OrElse inegalitaire = "OUI" OrElse inegalitaire = "1" OrElse inegalitaire = "X" Then
            values(FIELD_MAPPING("PartageInegalitaire")) = "/Oui"
        End If

        ' --- Section B : Ayants droit (max 17 lignes) ---
        Dim lineNum As Integer = 1
        For Each ayant In _data.AyantsDroit.Take(17)
            Dim fields As Dictionary(Of String, String) = GetAyantFieldIds(lineNum)
            Dim bdo = ayant.BDO
            Dim ident = ayant.Identite

            ' Role
            If bdo IsNot Nothing Then
                values(fields("Role")) = If(bdo.Role, "")
            End If

            ' Designation
            Dim designation As String
            If ident IsNot Nothing Then
                If String.Equals(ident.Type, "moral", StringComparison.OrdinalIgnoreCase) Then
                    designation = If(ident.Designation, "")
                Else
                    designation = GetDisplayCivilName(ident)
                End If

                ' Ajouter la société de gestion entre parenthèses pour les NON-SACEM
                Dim societeGestion As String = If(ident.SocieteGestion, "").Trim().ToUpper()
                Dim isNonSACEM As Boolean = Not String.IsNullOrEmpty(societeGestion) AndAlso societeGestion <> "SACEM"

                ' Détecter part inédite : A/C sans éditeur sur son lettrage
                Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
                Dim isAC As Boolean = (role = "A" OrElse role = "C" OrElse role = "AC" OrElse role = "AR" OrElse role = "AD")
                Dim isInedite As Boolean = False
                If isAC Then
                    Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
                    If Not String.IsNullOrEmpty(lettrage) Then
                        Dim hasEditeur As Boolean = _data.AyantsDroit.Any(Function(e)
                            Return e.BDO.Role = "E" AndAlso
                                   If(e.BDO.Lettrage, "").Trim().ToUpper() = lettrage
                        End Function)
                        isInedite = Not hasEditeur
                    End If
                End If

                ' Construire le suffixe parenthèses
                If isNonSACEM AndAlso isInedite Then
                    designation = $"{designation} ({ident.SocieteGestion.Trim()} – part inédite)"
                ElseIf isNonSACEM Then
                    designation = $"{designation} ({ident.SocieteGestion.Trim()})"
                ElseIf isInedite Then
                    designation = $"{designation} (part inédite)"
                End If

                values(fields("Designation")) = designation
            End If

            ' COAD/IPI (dans BDO, pas Identite)
            If bdo IsNot Nothing Then
                values(fields("COAD_IPI")) = If(bdo.COAD_IPI, "")
            End If

            ' Lettrage
            If bdo IsNot Nothing Then
                values(fields("Lettrage")) = If(bdo.Lettrage, "")

                ' PH (Clé Phono)
                If Not String.IsNullOrEmpty(bdo.PH) Then
                    Dim phVal As Double
                    If Double.TryParse(bdo.PH.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, phVal) Then
                        values(fields("PH")) = phVal.ToString("F2").Replace(".", ",")
                    Else
                        values(fields("PH")) = bdo.PH
                    End If
                End If
            End If

            lineNum += 1
        Next

        ' --- Page 2 : Date "Fait le" ---
        Dim faitle As String = If(_data.Faitle, "").Replace("""", "")
        If Not String.IsNullOrEmpty(faitle) Then
            Dim parts = faitle.Split("/"c)
            If parts.Length = 3 Then
                values(FIELD_MAPPING("FJJ")) = parts(0)
                values(FIELD_MAPPING("FMM")) = parts(1)
                values(FIELD_MAPPING("FAAAA")) = parts(2)
            End If
        End If

        ' Commentaire auto-généré
        values(FIELD_MAPPING("Commentaire")) = BuildBdoComment()

        Return values
    End Function

    ''' <summary>
    ''' Retourne les IDs des champs pour une ligne d'ayant droit (1-17)
    ''' Formule: base = 49 + (lineNum - 1) * 6
    ''' </summary>
    Private Function GetAyantFieldIds(lineNum As Integer) As Dictionary(Of String, String)
        Dim base As Integer = 49 + (lineNum - 1) * 6
        Return New Dictionary(Of String, String) From {
            {"Role", base.ToString()},
            {"Designation", (base + 1).ToString()},
            {"Lettrage", (base + 2).ToString()},
            {"DateContrat", (base + 3).ToString()},
            {"COAD_IPI", (base + 4).ToString()},
            {"PH", (base + 5).ToString()}
        }
    End Function

    ''' <summary>
    ''' Affichage civil pour le BDO : "Prénom NOM"
    ''' </summary>
    Private Function GetDisplayCivilName(ident As Identite) As String
        Dim prenom As String = TitleKeepHyphens(If(ident.Prenom, ""))
        Dim nom As String = If(ident.Nom, "").Trim().ToUpper()

        If Not String.IsNullOrEmpty(prenom) AndAlso Not String.IsNullOrEmpty(nom) Then
            Return $"{prenom} {nom}".Trim()
        ElseIf Not String.IsNullOrEmpty(nom) Then
            Return nom
        ElseIf Not String.IsNullOrEmpty(prenom) Then
            Return prenom
        Else
            Return If(ident.Designation, "")
        End If
    End Function

    ''' <summary>
    ''' Met en majuscule la première lettre de chaque mot, préserve les tirets
    ''' </summary>
    Private Function TitleKeepHyphens(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        s = s.Trim()

        Dim parts As New List(Of String)
        For Each tok In s.Replace("-", " - ").Split(" "c)
            If tok = "-" Then
                parts.Add("-")
            ElseIf tok.Length > 0 Then
                parts.Add(tok.Substring(0, 1).ToUpper() & tok.Substring(1).ToLower())
            End If
        Next

        Return String.Join("", parts).Replace(" - ", "-")
    End Function

    ''' <summary>
    ''' Construit le commentaire BDO.
    ''' Gère 3 cas :
    '''   1. Dépôt partiel (certains Signataire=FALSE) → "Le présent dépôt porte sur X% (part de ...)"
    '''   2. Dépôt complet avec non-SACEM → "Le présent dépôt porte sur X%... Les autres Y% reviennent à..."
    '''   3. Dépôt complet 100% SACEM → chaîne vide
    ''' </summary>
    Private Function BuildBdoComment() As String

        ' ── Détecter dépôt partiel ──────────────────────────────────────────
        Dim estPartiel As Boolean = _data.AyantsDroit.Any(Function(a) Not a.BDO.Signataire)

        If estPartiel Then
            Return BuildCommentairePartiel()
        End If

        ' ── Dépôt complet : logique non-SACEM existante ─────────────────────
        Dim personsSacem As New List(Of String)
        Dim editorsSacem As New HashSet(Of String)
        Dim nonSacemEditors As New List(Of Tuple(Of String, String)) ' (publisher, lettrage)
        Dim letterToPersons As New Dictionary(Of String, List(Of String))
        Dim totalNonSacem As Double = 0.0

        For Each ayant In _data.AyantsDroit
            Dim bdo = ayant.BDO
            Dim ident = ayant.Identite
            If bdo Is Nothing OrElse ident Is Nothing Then Continue For

            Dim role As String = If(bdo.Role, "").Trim().ToUpper()
            Dim lettrage As String = If(bdo.Lettrage, "").Trim()
            Dim estSacem As Boolean = IsSacem(ident)

            If role = "A" OrElse role = "AD" OrElse role = "C" OrElse role = "AR" Then
                Dim nm As String = GetDisplayCivilName(ident)
                If Not String.IsNullOrEmpty(nm) Then
                    If Not letterToPersons.ContainsKey(lettrage) Then
                        letterToPersons(lettrage) = New List(Of String)
                    End If
                    letterToPersons(lettrage).Add(nm)
                    If estSacem Then
                        personsSacem.Add(nm)
                    Else
                        totalNonSacem += ParseDouble(bdo.PH)
                    End If
                End If
            ElseIf role = "E" Then
                Dim pub As String = If(ident.Designation, "").Trim()
                If estSacem AndAlso Not String.IsNullOrEmpty(pub) Then
                    editorsSacem.Add(pub)
                ElseIf Not estSacem Then
                    totalNonSacem += ParseDouble(bdo.PH)
                    If Not String.IsNullOrEmpty(pub) Then
                        nonSacemEditors.Add(Tuple.Create(pub, lettrage))
                    End If
                End If
            End If
        Next

        personsSacem = personsSacem.Distinct().OrderBy(Function(x) x.ToLower()).ToList()
        Dim editorsList As List(Of String) = editorsSacem.OrderBy(Function(x) x.ToLower()).ToList()
        Dim partSacem As Double = Math.Max(0.0, 100.0 - totalNonSacem)
        Dim auteursStr As String = If(personsSacem.Any(), JoinListFr(personsSacem), "les ayants droits SACEM")
        Dim editeursStr As String = If(editorsList.Any(), JoinListFr(editorsList), "les éditeurs SACEM")

        If totalNonSacem > 0 Then
            Dim agg As New Dictionary(Of String, HashSet(Of String))
            For Each item In nonSacemEditors
                Dim pub As String = If(String.IsNullOrEmpty(item.Item1), "un éditeur non-SACEM", item.Item1)
                If Not agg.ContainsKey(pub) Then agg(pub) = New HashSet(Of String)
                If Not String.IsNullOrEmpty(item.Item2) AndAlso letterToPersons.ContainsKey(item.Item2) Then
                    For Each p In letterToPersons(item.Item2)
                        agg(pub).Add(p)
                    Next
                End If
            Next

            Dim segs As New List(Of String)
            For Each pub In agg.Keys.OrderBy(Function(x) x.ToLower())
                Dim persons = agg(pub).OrderBy(Function(x) x.ToLower()).ToList()
                segs.Add(If(persons.Any(), $"{pub} éditeur de {JoinListFr(persons)}", pub))
            Next

            Dim nonSacemPhrase As String = If(segs.Any(), JoinListFr(segs), "des ayants droits non membres de la SACEM")
            Return $"Le présent dépôt porte sur {FormatPct(partSacem)} de l'oeuvre " &
                   $"(parts de {auteursStr} et leurs éditeurs {editeursStr}). " &
                   $"Les autres {FormatPct(totalNonSacem)} reviennent à {nonSacemPhrase}, " &
                   "membres d'une société de gestion collective étrangère."
        End If

        ' Tous SACEM 100% : pas de commentaire
        Return ""
    End Function

    ''' <summary>
    ''' Construit le commentaire pour un dépôt partiel.
    ''' Calcule le % total des signataires et liste leurs noms (créateurs + éditeurs).
    ''' </summary>
    Private Function BuildCommentairePartiel() As String
        Dim totalPart As Double = 0.0
        Dim noms As New List(Of String)

        For Each ayant In _data.AyantsDroit
            If Not ayant.BDO.Signataire Then Continue For

            Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
            Dim ph As Double = ParseDouble(ayant.BDO.PH)
            totalPart += ph

            ' Collecter les noms des signataires (créateurs et éditeurs)
            If role = "A" OrElse role = "C" OrElse role = "AC" OrElse role = "AR" OrElse role = "AD" Then
                Dim nm As String = GetDisplayCivilName(ayant.Identite)
                If Not String.IsNullOrEmpty(nm) AndAlso Not noms.Contains(nm) Then
                    noms.Add(nm)
                End If
            ElseIf role = "E" Then
                Dim pub As String = If(ayant.Identite.Designation, "").Trim()
                If Not String.IsNullOrEmpty(pub) AndAlso Not noms.Contains(pub) Then
                    noms.Add(pub)
                End If
            End If
        Next

        Dim nomsStr As String = If(noms.Any(), JoinListFr(noms), "les signataires")
        Return $"Le présent dépôt porte sur {FormatPct(totalPart)} de l'oeuvre (part de {nomsStr})"
    End Function

    ''' <summary>
    ''' Vérifie si un ayant droit est membre de la SACEM
    ''' </summary>
    Private Function IsSacem(ident As Identite) As Boolean
        If ident Is Nothing Then Return False
        Dim sg As String = If(ident.SocieteGestion, "").Trim().ToUpper()
        Return sg = "SACEM"
    End Function

    ''' <summary>
    ''' Parse un nombre en Double
    ''' </summary>
    Private Function ParseDouble(s As String) As Double
        If String.IsNullOrEmpty(s) Then Return 0.0
        Dim val As Double
        If Double.TryParse(s.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, val) Then
            Return val
        End If
        Return 0.0
    End Function

    ''' <summary>
    ''' Formate un pourcentage en français
    ''' </summary>
    Private Function FormatPct(v As Double) As String
        Return v.ToString("F2").Replace(".", ",") & "%"
    End Function

    ''' <summary>
    ''' Joint une liste avec des virgules et "et" pour le dernier
    ''' </summary>
    Private Function JoinListFr(items As List(Of String)) As String
        items = items.Where(Function(i) Not String.IsNullOrEmpty(i)).ToList()
        If Not items.Any() Then Return ""
        If items.Count = 1 Then Return items(0)
        Return String.Join(", ", items.Take(items.Count - 1)) & " et " & items.Last()
    End Function

    ''' <summary>
    ''' Écrit les valeurs des champs dans un fichier JSON
    ''' </summary>
    Private Sub WriteFieldValuesJson(values As Dictionary(Of String, String), jsonPath As String)
        Dim sb As New StringBuilder()
        sb.AppendLine("{")

        Dim first As Boolean = True
        For Each kvp In values
            If Not first Then sb.AppendLine(",")
            first = False

            Dim escapedValue As String = kvp.Value.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "").Replace(vbLf, " ")
            sb.Append($"  ""{kvp.Key}"": ""{escapedValue}""")
        Next

        sb.AppendLine()
        sb.AppendLine("}")

        File.WriteAllText(jsonPath, sb.ToString(), Encoding.UTF8)
    End Sub

    ''' <summary>
    ''' Appelle le script Python pour remplir le PDF
    ''' </summary>
    Private Function FillPdfWithPython(templatePath As String, jsonPath As String, outputPath As String) As Boolean
        Try
            ' Chemin du script Python FillBDOFromJson.py (dans Scripts/FillBDO/)
            Dim scriptPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts", "FillBDO", "FillBDOFromJson.py")

            ' Si le script n'existe pas, le créer
            If Not File.Exists(scriptPath) Then
                CreateFillBDOScript(scriptPath)
            End If

            Dim psi As New ProcessStartInfo()
            psi.FileName = "python"
            psi.Arguments = $"""{scriptPath}"" ""{templatePath}"" ""{jsonPath}"" ""{outputPath}"""
            psi.UseShellExecute = False
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True

            Using process As Process = Process.Start(psi)
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim errors As String = process.StandardError.ReadToEnd()
                process.WaitForExit()

                If Not String.IsNullOrEmpty(output) Then
                    _log.Add($"Python: {output.Trim()}")
                End If

                If process.ExitCode <> 0 Then
                    _log.Add($"✗ Erreur Python: {errors}")
                    Return False
                End If

                Return True
            End Using

        Catch ex As Exception
            _log.Add($"✗ Erreur appel Python: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Crée le script Python de remplissage PDF
    ''' </summary>
    Private Sub CreateFillBDOScript(scriptPath As String)
        Dim script As String = "#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import json
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, NumberObject, ArrayObject

# Champ Commentaire = ""10""
COMMENT_FIELD = ""10""

# Paramètres d'auto-ajustement pour le champ Commentaire
# Largeur approximative du champ en points (à ajuster selon le PDF)
COMMENT_FIELD_WIDTH = 500
COMMENT_FIELD_HEIGHT = 60

# Tailles de police min/max
FONT_SIZE_MAX = 10
FONT_SIZE_MIN = 5

def calculate_font_size(text, max_width, max_height, max_font=FONT_SIZE_MAX, min_font=FONT_SIZE_MIN):
    """"""
    Calcule la taille de police optimale pour que le texte tienne dans le champ.
    Approximation basée sur la longueur du texte.
    """"""
    if not text:
        return max_font
    
    text_len = len(text)
    
    # Estimation : environ 2.5 caractères par point de largeur en taille 10
    # Plus le texte est long, plus on réduit la taille
    chars_per_line_at_10pt = max_width / 4.5  # ~4.5 points par caractère en taille 10
    lines_at_10pt = max_height / 12  # ~12 points par ligne en taille 10
    max_chars_at_10pt = chars_per_line_at_10pt * lines_at_10pt
    
    if text_len <= max_chars_at_10pt * 0.5:
        return max_font
    elif text_len <= max_chars_at_10pt * 0.7:
        return max(min_font, max_font - 1)
    elif text_len <= max_chars_at_10pt:
        return max(min_font, max_font - 2)
    elif text_len <= max_chars_at_10pt * 1.3:
        return max(min_font, max_font - 3)
    elif text_len <= max_chars_at_10pt * 1.6:
        return max(min_font, max_font - 4)
    else:
        return min_font

def set_field_font_size(writer, field_name, font_size):
    """"""
    Définit la taille de police pour un champ spécifique.
    """"""
    try:
        for page in writer.pages:
            if '/Annots' in page:
                for annot in page['/Annots']:
                    annot_obj = annot.get_object()
                    if annot_obj.get('/T') == field_name:
                        # Créer l'apparence par défaut avec la nouvelle taille
                        da = f'/Helv {font_size} Tf 0 g'
                        annot_obj[NameObject('/DA')] = TextStringObject(da)
                        return True
    except Exception as e:
        print(f'Erreur set_field_font_size: {e}')
    return False

def main():
    if len(sys.argv) < 4:
        print('Usage: python FillBDOFromJson.py <template.pdf> <values.json> <output.pdf>')
        sys.exit(1)
    
    template_path = sys.argv[1]
    json_path = sys.argv[2]
    output_path = sys.argv[3]
    
    with open(json_path, 'r', encoding='utf-8') as f:
        field_values = json.load(f)
    
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)
    
    # Calculer la taille de police pour le champ Commentaire
    comment_text = field_values.get(COMMENT_FIELD, '')
    if comment_text:
        font_size = calculate_font_size(comment_text, COMMENT_FIELD_WIDTH, COMMENT_FIELD_HEIGHT)
        print(f'Commentaire: {len(comment_text)} chars -> police {font_size}pt')
        set_field_font_size(writer, COMMENT_FIELD, font_size)
    
    # Remplir les champs
    for page in writer.pages:
        writer.update_page_form_field_values(page, field_values)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    print(f'PDF rempli: {output_path}')

if __name__ == '__main__':
    main()
"
        ' Créer le dossier si nécessaire
        Dim scriptDir As String = Path.GetDirectoryName(scriptPath)
        If Not Directory.Exists(scriptDir) Then
            Directory.CreateDirectory(scriptDir)
        End If

        File.WriteAllText(scriptPath, script, Encoding.UTF8)
        _log.Add($"Script Python créé: {scriptPath}")
    End Sub

    ''' <summary>
    ''' Génère le BDO avec uniquement les données (sans appel Python - pour test)
    ''' Retourne le dictionnaire des valeurs
    ''' </summary>
    Public Function GenerateFieldValuesOnly() As Dictionary(Of String, String)
        _log.Clear()
        _log.Add("=== GÉNÉRATION VALEURS BDO ===")
        Dim values = GenerateFieldValues()
        _log.Add($"Champs générés: {values.Count}")
        Return values
    End Function
End Class
