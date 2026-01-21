Imports System.Linq
Imports Microsoft.Office.Interop.Word

''' <summary>
''' Export de documents Word (DOCX) vers PDF
''' Utilise Microsoft.Office.Interop.Word (nécessite Microsoft Word installé)
''' </summary>
Public Class PdfExporter
    Private _log As New List(Of String)

    ''' <summary>
    ''' Log des opérations d'export
    ''' </summary>
    Public ReadOnly Property ExportLog As List(Of String)
        Get
            Return _log
        End Get
    End Property

    ''' <summary>
    ''' Exporte un document DOCX vers PDF
    ''' </summary>
    ''' <param name="docxPath">Chemin du fichier DOCX source</param>
    ''' <param name="pdfPath">Chemin du fichier PDF de destination</param>
    ''' <returns>True si succès, False sinon</returns>
    Public Function ExportToPdf(docxPath As String, pdfPath As String) As Boolean
        Dim wordApp As Application = Nothing
        Dim wordDoc As Document = Nothing

        Try
            _log.Clear()
            _log.Add("Conversion DOCX → PDF...")

            ' Vérifications
            If Not System.IO.File.Exists(docxPath) Then
                Throw New System.IO.FileNotFoundException($"Fichier DOCX introuvable: {docxPath}")
            End If

            ' Créer le dossier de destination si nécessaire
            Dim outputDir As String = System.IO.Path.GetDirectoryName(pdfPath)
            If Not String.IsNullOrEmpty(outputDir) AndAlso Not System.IO.Directory.Exists(outputDir) Then
                System.IO.Directory.CreateDirectory(outputDir)
            End If

            ' Initialiser Word
            _log.Add("  Démarrage de Microsoft Word...")
            wordApp = New Application()
            wordApp.Visible = False
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone

            ' Ouvrir le document
            _log.Add($"  Ouverture du document: {System.IO.Path.GetFileName(docxPath)}")
            wordDoc = wordApp.Documents.Open(docxPath, [ReadOnly]:=True)

            ' Exporter en PDF
            _log.Add($"  Export vers: {System.IO.Path.GetFileName(pdfPath)}")
            wordDoc.ExportAsFixedFormat(
                OutputFileName:=pdfPath,
                ExportFormat:=WdExportFormat.wdExportFormatPDF,
                OpenAfterExport:=False,
                OptimizeFor:=WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range:=WdExportRange.wdExportAllDocument,
                Item:=WdExportItem.wdExportDocumentContent,
                IncludeDocProps:=True,
                KeepIRM:=True,
                CreateBookmarks:=WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                DocStructureTags:=True,
                BitmapMissingFonts:=True,
                UseISO19005_1:=False
            )

            _log.Add("✓ Conversion réussie")
            Return True

        Catch ex As Exception
            _log.Add($"✗ Erreur conversion PDF: {ex.Message}")
            
            If ex.InnerException IsNot Nothing Then
                _log.Add($"  Détail: {ex.InnerException.Message}")
            End If

            ' Messages d'erreur spécifiques
            If ex.Message.Contains("0x800A03EC") Then
                _log.Add("  → Microsoft Word n'est pas installé ou n'est pas accessible")
            ElseIf ex.Message.Contains("RPC") Then
                _log.Add("  → Problème de communication avec Microsoft Word")
                _log.Add("  → Essayez de fermer toutes les instances de Word")
            End If

            Return False

        Finally
            ' Nettoyage COM
            Try
                If wordDoc IsNot Nothing Then
                    wordDoc.Close(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc)
                    wordDoc = Nothing
                End If

                If wordApp IsNot Nothing Then
                    wordApp.Quit(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                    wordApp = Nothing
                End If

                ' Forcer le garbage collector pour libérer les ressources COM
                GC.Collect()
                GC.WaitForPendingFinalizers()

            Catch cleanupEx As Exception
                _log.Add($"  Avertissement nettoyage: {cleanupEx.Message}")
            End Try
        End Try
    End Function

    ''' <summary>
    ''' Vérifie si Microsoft Word est installé et accessible
    ''' </summary>
    ''' <returns>True si Word est disponible, False sinon</returns>
    Public Shared Function IsWordInstalled() As Boolean
        Dim wordApp As Application = Nothing

        Try
            wordApp = New Application()
            Return True

        Catch ex As Exception
            Return False

        Finally
            If wordApp IsNot Nothing Then
                Try
                    wordApp.Quit(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                Catch
                    ' Ignorer les erreurs de nettoyage
                End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Conversion par lot de plusieurs fichiers DOCX vers PDF
    ''' </summary>
    ''' <param name="docxPaths">Liste des chemins DOCX</param>
    ''' <param name="outputDirectory">Dossier de sortie pour les PDFs</param>
    ''' <returns>Dictionnaire avec le statut de chaque conversion</returns>
    Public Function BatchExportToPdf(docxPaths As List(Of String), outputDirectory As String) As Dictionary(Of String, Boolean)
        Dim results As New Dictionary(Of String, Boolean)

        _log.Clear()
        _log.Add($"=== CONVERSION PAR LOT ({docxPaths.Count} fichiers) ===")

        For Each docxPath In docxPaths
            Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(docxPath)
            Dim pdfPath As String = System.IO.Path.Combine(outputDirectory, fileName & ".pdf")

            _log.Add($"Conversion: {fileName}")
            Dim success As Boolean = ExportToPdf(docxPath, pdfPath)
            results(docxPath) = success
        Next

        Dim successCount As Integer = results.Values.Where(Function(v) v = True).Count()
        _log.Add($"=== CONVERSION TERMINÉE: {successCount}/{docxPaths.Count} réussies ===")

        Return results
    End Function
End Class
