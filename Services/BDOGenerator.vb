Imports System.IO
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

''' <summary>
''' Générateur de Bulletins de Déclaration d'Œuvre (BDO) SACEM
''' Approche: Génération DOCX puis conversion PDF
''' Alternative à l'approche PyMuPDF du script Python
''' </summary>
Public Class BDOGenerator
    Private ReadOnly _data As SACEMData
    Private _log As New List(Of String)

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
    ''' Génère un BDO complet (DOCX puis PDF)
    ''' </summary>
    Public Function Generate(bdoTemplatePath As String, outputPath As String) As Boolean
        Try
            _log.Clear()
            _log.Add("=== GÉNÉRATION BDO SACEM ===")
            _log.Add($"Titre: {_data.Titre}")
            _log.Add($"Interprète: {_data.Interprete}")

            If Not File.Exists(bdoTemplatePath) Then
                Throw New FileNotFoundException($"Template BDO introuvable: {bdoTemplatePath}")
            End If

            ' 1. Créer le DOCX temporaire
            Dim tempDocx As String = Path.Combine(Path.GetTempPath(), $"BDO_{Guid.NewGuid()}.docx")
            _log.Add("Génération du DOCX...")

            If Not GenerateBDODocx(bdoTemplatePath, tempDocx) Then
                Throw New Exception("Échec de la génération du DOCX")
            End If

            _log.Add("✓ DOCX généré")

            ' 2. Convertir en PDF
            _log.Add("Conversion en PDF...")
            Dim pdfExporter As New PdfExporter()
            Dim pdfSuccess As Boolean = pdfExporter.ExportToPdf(tempDocx, outputPath)

            For Each logEntry In pdfExporter.ExportLog
                _log.Add(logEntry)
            Next

            ' 3. Nettoyer
            Try
                If File.Exists(tempDocx) Then File.Delete(tempDocx)
            Catch
            End Try

            If pdfSuccess Then
                _log.Add("✓ BDO généré avec succès")
                _log.Add($"=== GÉNÉRATION TERMINÉE ===")
                Return True
            Else
                _log.Add("✗ Échec de la conversion PDF")
                Return False
            End If

        Catch ex As Exception
            _log.Add($"✗ ERREUR: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Génère le DOCX du BDO avec balises indexées
    ''' </summary>
    Private Function GenerateBDODocx(templatePath As String, outputPath As String) As Boolean
        Try
            File.Copy(templatePath, outputPath, overwrite:=True)

            Dim balisesGenerator As New BalisesGenerator(_data)
            Dim allBalises As Dictionary(Of String, String) = balisesGenerator.GenerateAllBalises()

            _log.Add($"  {allBalises.Count} balises générées")
            _log.Add($"  {_data.AyantsDroit.Count} ayants droit à traiter")

            Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(outputPath, True)
                If wordDoc.MainDocumentPart Is Nothing Then
                    Throw New InvalidDataException("Document Word invalide")
                End If

                Dim body As Body = wordDoc.MainDocumentPart.Document.Body
                ReplaceAllBalisesInBDO(body, allBalises)
                wordDoc.MainDocumentPart.Document.Save()
            End Using

            _log.Add($"  {_data.AyantsDroit.Count} ayants droit insérés")
            Return True

        Catch ex As Exception
            _log.Add($"  ✗ Erreur génération DOCX: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Remplace toutes les balises dans le BDO
    ''' </summary>
    Private Sub ReplaceAllBalisesInBDO(body As Body, balises As Dictionary(Of String, String))
        Try
            Dim paragraphs As List(Of Paragraph) = body.Descendants(Of Paragraph)().ToList()

            For Each paragraph In paragraphs
                Dim paragraphText As String = GetParagraphText(paragraph)

                ' Balises simples [xxx]
                Dim simpleMatches As MatchCollection = Regex.Matches(paragraphText, "\[([^\]]+)\]")
                For Each match As Match In simpleMatches
                    Dim key As String = match.Groups(1).Value
                    If balises.ContainsKey(key) Then
                        ReplaceInParagraph(paragraph, match.Value, balises(key))
                    End If
                Next

                ' Balises indexées {xxx}
                Dim indexedMatches As MatchCollection = Regex.Matches(paragraphText, "\{([^\}]+)\}")
                For Each match As Match In indexedMatches
                    Dim key As String = match.Groups(1).Value
                    If balises.ContainsKey(key) Then
                        ReplaceInParagraph(paragraph, match.Value, balises(key))
                    End If
                Next
            Next

            ' Traiter les tableaux
            For Each table In body.Descendants(Of Table)()
                ReplaceInTable(table, balises)
            Next

        Catch ex As Exception
            _log.Add($"  ✗ Erreur remplacement balises: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Remplace les balises dans un tableau
    ''' </summary>
    Private Sub ReplaceInTable(table As Table, balises As Dictionary(Of String, String))
        For Each row In table.Descendants(Of TableRow)()
            For Each cell In row.Descendants(Of TableCell)()
                For Each paragraph In cell.Descendants(Of Paragraph)()
                    Dim paragraphText As String = GetParagraphText(paragraph)

                    Dim simpleMatches As MatchCollection = Regex.Matches(paragraphText, "\[([^\]]+)\]")
                    For Each match As Match In simpleMatches
                        Dim key As String = match.Groups(1).Value
                        If balises.ContainsKey(key) Then
                            ReplaceInParagraph(paragraph, match.Value, balises(key))
                        End If
                    Next

                    Dim indexedMatches As MatchCollection = Regex.Matches(paragraphText, "\{([^\}]+)\}")
                    For Each match As Match In indexedMatches
                        Dim key As String = match.Groups(1).Value
                        If balises.ContainsKey(key) Then
                            ReplaceInParagraph(paragraph, match.Value, balises(key))
                        End If
                    Next
                Next
            Next
        Next
    End Sub

    ''' <summary>
    ''' Remplace du texte dans un paragraphe
    ''' </summary>
    Private Sub ReplaceInParagraph(paragraph As Paragraph, oldText As String, newText As String)
        Try
            Dim runs As List(Of Run) = paragraph.Descendants(Of Run)().ToList()
            Dim fullText As String = String.Join("", runs.Select(Function(r) r.InnerText))

            If Not fullText.Contains(oldText) Then Return

            Dim newFullText As String = fullText.Replace(oldText, newText)

            For Each run In runs
                run.Remove()
            Next

            Dim newRun As New Run()
            If runs.Count > 0 AndAlso runs(0).RunProperties IsNot Nothing Then
                newRun.RunProperties = CType(runs(0).RunProperties.CloneNode(True), RunProperties)
            End If

            Dim textElement As New Text(newFullText)
            textElement.Space = "preserve"
            newRun.Append(textElement)
            paragraph.Append(newRun)

        Catch ex As Exception
            Debug.WriteLine($"Erreur ReplaceInParagraph: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Récupère le texte d'un paragraphe
    ''' </summary>
    Private Function GetParagraphText(paragraph As Paragraph) As String
        Return String.Join("", paragraph.Descendants(Of Text)().Select(Function(t) t.Text))
    End Function

    ''' <summary>
    ''' Génère uniquement le DOCX sans conversion PDF
    ''' </summary>
    Public Function GenerateDocxOnly(bdoTemplatePath As String, outputPath As String) As Boolean
        Try
            _log.Clear()
            _log.Add("=== GÉNÉRATION BDO DOCX ===")

            If Not File.Exists(bdoTemplatePath) Then
                Throw New FileNotFoundException($"Template BDO introuvable: {bdoTemplatePath}")
            End If

            Dim success As Boolean = GenerateBDODocx(bdoTemplatePath, outputPath)

            If success Then
                _log.Add("✓ BDO DOCX généré avec succès")
            Else
                _log.Add("✗ Échec de la génération")
            End If

            Return success

        Catch ex As Exception
            _log.Add($"✗ ERREUR: {ex.Message}")
            Return False
        End Try
    End Function
End Class
