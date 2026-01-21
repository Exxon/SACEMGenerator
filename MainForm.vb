Imports System.IO
Imports System.Windows.Forms

''' <summary>
''' Interface graphique principale pour le générateur SACEM
''' Permet de générer BDO et Contrats à partir de fichiers JSON
''' </summary>
Public Class MainForm
    Inherits Form
    
    Private _currentJsonPath As String
    Private _currentData As SACEMData
    Private _paragraphReader As ParagraphTemplateReader

    ' Chemins de configuration
    Private _templatesDirectory As String = "Templates"
    Private _outputDirectory As String = "Output"
    Private _paragraphTemplatePath As String = "Templates\template_paragrahs.docx"

    ''' <summary>
    ''' Initialisation du formulaire
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        InitializePaths()
    End Sub

    ''' <summary>
    ''' Initialise les chemins par défaut
    ''' </summary>
    Private Sub InitializePaths()
        ' Créer les répertoires s'ils n'existent pas
        If Not Directory.Exists(_templatesDirectory) Then
            Directory.CreateDirectory(_templatesDirectory)
        End If

        If Not Directory.Exists(_outputDirectory) Then
            Directory.CreateDirectory(_outputDirectory)
        End If

        ' Afficher les chemins dans l'interface
        txtTemplatesPath.Text = Path.GetFullPath(_templatesDirectory)
        txtOutputPath.Text = Path.GetFullPath(_outputDirectory)
        txtParagraphTemplate.Text = Path.GetFullPath(_paragraphTemplatePath)

        ' Désactiver les boutons de génération au démarrage
        btnGenerateBDO.Enabled = False
        btnGenerateContracts.Enabled = False
        btnGenerateAll.Enabled = False
    End Sub

    ''' <summary>
    ''' Bouton "Sélectionner JSON"
    ''' </summary>
    Private Sub btnSelectJson_Click(sender As Object, e As EventArgs) Handles btnSelectJson.Click
        Using dialog As New OpenFileDialog()
            dialog.Title = "Sélectionner le fichier JSON SACEM"
            dialog.Filter = "Fichiers JSON (*.json)|*.json|Tous les fichiers (*.*)|*.*"
            dialog.FilterIndex = 1

            If dialog.ShowDialog() = DialogResult.OK Then
                _currentJsonPath = dialog.FileName
                txtJsonPath.Text = _currentJsonPath

                ' Charger et valider le JSON
                LoadAndValidateJson()
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Charge et valide le fichier JSON
    ''' </summary>
    Private Sub LoadAndValidateJson()
        Try
            txtLog.AppendText($"{vbCrLf}=== CHARGEMENT JSON ==={vbCrLf}")
            txtLog.AppendText($"Fichier: {Path.GetFileName(_currentJsonPath)}{vbCrLf}")

            ' Charger les données
            _currentData = SACEMJsonReader.LoadFromFile(_currentJsonPath)

            If _currentData Is Nothing Then
                Throw New Exception("Échec du chargement du JSON")
            End If

            ' Valider la structure
            Dim validation = SACEMJsonReader.ValidateStructure(_currentData)
            If Not validation.IsValid Then
                txtLog.AppendText($"⚠ ATTENTION: {validation.ErrorMessage}{vbCrLf}")
            Else
                txtLog.AppendText($"✓ Structure JSON valide{vbCrLf}")
            End If

            ' Générer un rapport de structure
            Dim report As String = SACEMJsonReader.GenerateStructureReport(_currentData)
            txtLog.AppendText($"{report}{vbCrLf}")

            ' Afficher les informations principales
            lblTitre.Text = $"Titre: {_currentData.Titre}"
            lblInterprete.Text = $"Interprète: {_currentData.Interprete}"
            lblAyantsDroit.Text = $"Ayants droit: {_currentData.AyantsDroit.Count}"

            ' Charger les templates de paragraphes
            LoadParagraphTemplates()

            ' Activer les boutons de génération
            btnGenerateBDO.Enabled = True
            btnGenerateContracts.Enabled = True
            btnGenerateAll.Enabled = True

            txtLog.AppendText($"✓ JSON chargé et validé avec succès{vbCrLf}")

        Catch ex As Exception
            txtLog.AppendText($"✗ ERREUR: {ex.Message}{vbCrLf}")
            MessageBox.Show($"Erreur de chargement du JSON:{vbCrLf}{ex.Message}", 
                           "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

            btnGenerateBDO.Enabled = False
            btnGenerateContracts.Enabled = False
            btnGenerateAll.Enabled = False
        End Try
    End Sub

    ''' <summary>
    ''' Charge les templates de paragraphes
    ''' </summary>
    Private Sub LoadParagraphTemplates()
        Try
            txtLog.AppendText($"{vbCrLf}Chargement des templates de paragraphes...{vbCrLf}")

            If Not File.Exists(_paragraphTemplatePath) Then
                txtLog.AppendText($"⚠ ATTENTION: Template de paragraphes introuvable{vbCrLf}")
                txtLog.AppendText($"  Chemin: {_paragraphTemplatePath}{vbCrLf}")
                txtLog.AppendText($"  Les superbalises ne fonctionneront pas correctement{vbCrLf}")
                Return
            End If

            _paragraphReader = New ParagraphTemplateReader(_paragraphTemplatePath)
            _paragraphReader.LoadTemplates()

            txtLog.AppendText($"✓ Templates de paragraphes chargés{vbCrLf}")

        Catch ex As Exception
            txtLog.AppendText($"⚠ Erreur chargement templates: {ex.Message}{vbCrLf}")
        End Try
    End Sub

    ''' <summary>
    ''' Bouton "Générer BDO"
    ''' </summary>
    Private Sub btnGenerateBDO_Click(sender As Object, e As EventArgs) Handles btnGenerateBDO.Click
        Try
            txtLog.AppendText($"{vbCrLf}=== GÉNÉRATION BDO ==={vbCrLf}")

            ' Vérifier le template BDO
            Dim bdoTemplatePath As String = Path.Combine(_templatesDirectory, "BDO_template.docx")
            If Not File.Exists(bdoTemplatePath) Then
                MessageBox.Show($"Template BDO introuvable:{vbCrLf}{bdoTemplatePath}", 
                               "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Nom du fichier de sortie
            Dim outputFileName As String = $"BDO_{_currentData.Titre}_{_currentData.Interprete}.pdf"
            outputFileName = CleanFileName(outputFileName)
            Dim outputPath As String = Path.Combine(_outputDirectory, outputFileName)

            ' Générer le BDO
            Dim generator As New BDOGenerator(_currentData)
            Dim success As Boolean = generator.Generate(bdoTemplatePath, outputPath)

            ' Afficher les logs
            For Each logEntry In generator.GenerationLog
                txtLog.AppendText($"{logEntry}{vbCrLf}")
            Next

            If success Then
                MessageBox.Show($"BDO généré avec succès:{vbCrLf}{outputPath}", 
                               "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Échec de la génération du BDO. Consultez les logs.", 
                               "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            txtLog.AppendText($"✗ ERREUR: {ex.Message}{vbCrLf}")
            MessageBox.Show($"Erreur:{vbCrLf}{ex.Message}", 
                           "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Bouton "Générer Contrats"
    ''' </summary>
    Private Sub btnGenerateContracts_Click(sender As Object, e As EventArgs) Handles btnGenerateContracts.Click
        Try
            txtLog.AppendText($"{vbCrLf}=== GÉNÉRATION CONTRATS ==={vbCrLf}")

            ' Vérifier les templates
            Dim templates As New Dictionary(Of String, String) From {
                {"CCDAA", Path.Combine(_templatesDirectory, "CCDAA_template.docx")},
                {"CCEOM", Path.Combine(_templatesDirectory, "CCEOM_template_univ.docx")},
                {"COED", Path.Combine(_templatesDirectory, "COED_template_univ.docx")}
            }

            Dim missingTemplates As New List(Of String)
            For Each kvp In templates
                If Not File.Exists(kvp.Value) Then
                    missingTemplates.Add(kvp.Key)
                End If
            Next

            If missingTemplates.Count > 0 Then
                Dim message As String = $"Templates manquants:{vbCrLf}{String.Join(", ", missingTemplates)}"
                MessageBox.Show(message, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ' Générer chaque type de contrat
            Dim successCount As Integer = 0
            Dim generator As New ContractGenerator(_currentData, _paragraphReader)

            For Each kvp In templates
                Dim contractType As String = kvp.Key
                Dim templatePath As String = kvp.Value

                If Not File.Exists(templatePath) Then
                    txtLog.AppendText($"⊗ {contractType}: Template manquant{vbCrLf}")
                    Continue For
                End If

                ' Nom du fichier de sortie
                Dim outputFileName As String = $"{contractType}_{_currentData.Titre}_{_currentData.Interprete}.docx"
                outputFileName = CleanFileName(outputFileName)
                Dim outputPath As String = Path.Combine(_outputDirectory, outputFileName)

                ' Générer le contrat
                Dim success As Boolean = generator.Generate(templatePath, outputPath, contractType)

                ' Afficher les logs
                For Each logEntry In generator.GenerationLog
                    txtLog.AppendText($"{logEntry}{vbCrLf}")
                Next

                If success Then
                    successCount += 1
                End If
            Next

            ' Message final
            Dim totalContracts As Integer = templates.Count
            If successCount = totalContracts Then
                MessageBox.Show($"Tous les contrats générés avec succès ({successCount}/{totalContracts})", 
                               "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf successCount > 0 Then
                MessageBox.Show($"Certains contrats générés ({successCount}/{totalContracts}). Consultez les logs.", 
                               "Succès partiel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                MessageBox.Show("Échec de la génération des contrats. Consultez les logs.", 
                               "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            txtLog.AppendText($"✗ ERREUR: {ex.Message}{vbCrLf}")
            MessageBox.Show($"Erreur:{vbCrLf}{ex.Message}", 
                           "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Bouton "Tout générer"
    ''' </summary>
    Private Sub btnGenerateAll_Click(sender As Object, e As EventArgs) Handles btnGenerateAll.Click
        Try
            txtLog.AppendText($"{vbCrLf}=== GÉNÉRATION COMPLÈTE (BDO + CONTRATS) ==={vbCrLf}")

            ' Générer le BDO
            btnGenerateBDO_Click(sender, e)

            ' Générer les contrats
            btnGenerateContracts_Click(sender, e)

            txtLog.AppendText($"{vbCrLf}=== GÉNÉRATION COMPLÈTE TERMINÉE ==={vbCrLf}")

        Catch ex As Exception
            txtLog.AppendText($"✗ ERREUR: {ex.Message}{vbCrLf}")
            MessageBox.Show($"Erreur:{vbCrLf}{ex.Message}", 
                           "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Bouton "Effacer les logs"
    ''' </summary>
    Private Sub btnClearLog_Click(sender As Object, e As EventArgs) Handles btnClearLog.Click
        txtLog.Clear()
        txtLog.AppendText("=== SACEM GENERATOR - VB.NET ===" & vbCrLf)
        txtLog.AppendText($"Version 1.0 - {DateTime.Now:yyyy-MM-dd}" & vbCrLf)
        txtLog.AppendText(vbCrLf)
    End Sub

    ''' <summary>
    ''' Bouton "Ouvrir dossier Output"
    ''' </summary>
    Private Sub btnOpenOutput_Click(sender As Object, e As EventArgs) Handles btnOpenOutput.Click
        Try
            If Directory.Exists(_outputDirectory) Then
                Process.Start("explorer.exe", _outputDirectory)
            Else
                MessageBox.Show("Le dossier Output n'existe pas encore.", 
                               "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Impossible d'ouvrir le dossier:{vbCrLf}{ex.Message}", 
                           "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Nettoie un nom de fichier des caractères invalides
    ''' </summary>
    Private Function CleanFileName(fileName As String) As String
        Dim invalidChars As Char() = Path.GetInvalidFileNameChars()
        For Each c As Char In invalidChars
            fileName = fileName.Replace(c, "_"c)
        Next
        Return fileName
    End Function

    ''' <summary>
    ''' Initialisation des composants du formulaire
    ''' (À compléter selon vos besoins d'interface)
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = "SACEM Generator - VB.NET"
        Me.Size = New Size(900, 700)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' GroupBox - Sélection JSON
        Dim grpJson As New GroupBox()
        grpJson.Text = "1. Sélection du fichier JSON"
        grpJson.Location = New Point(10, 10)
        grpJson.Size = New Size(860, 100)

        txtJsonPath = New TextBox()
        txtJsonPath.Location = New Point(10, 25)
        txtJsonPath.Size = New Size(700, 20)
        txtJsonPath.ReadOnly = True
        grpJson.Controls.Add(txtJsonPath)

        btnSelectJson = New Button()
        btnSelectJson.Text = "Parcourir..."
        btnSelectJson.Location = New Point(720, 23)
        btnSelectJson.Size = New Size(120, 25)
        grpJson.Controls.Add(btnSelectJson)

        lblTitre = New Label()
        lblTitre.Location = New Point(10, 55)
        lblTitre.Size = New Size(400, 20)
        lblTitre.Text = "Titre: -"
        grpJson.Controls.Add(lblTitre)

        lblInterprete = New Label()
        lblInterprete.Location = New Point(10, 75)
        lblInterprete.Size = New Size(400, 20)
        lblInterprete.Text = "Interprète: -"
        grpJson.Controls.Add(lblInterprete)

        lblAyantsDroit = New Label()
        lblAyantsDroit.Location = New Point(420, 55)
        lblAyantsDroit.Size = New Size(200, 20)
        lblAyantsDroit.Text = "Ayants droit: 0"
        grpJson.Controls.Add(lblAyantsDroit)

        Me.Controls.Add(grpJson)

        ' GroupBox - Configuration
        Dim grpConfig As New GroupBox()
        grpConfig.Text = "2. Configuration des chemins"
        grpConfig.Location = New Point(10, 120)
        grpConfig.Size = New Size(860, 100)

        Dim lblTemplates As New Label()
        lblTemplates.Text = "Dossier Templates:"
        lblTemplates.Location = New Point(10, 25)
        lblTemplates.Size = New Size(120, 20)
        grpConfig.Controls.Add(lblTemplates)

        txtTemplatesPath = New TextBox()
        txtTemplatesPath.Location = New Point(140, 23)
        txtTemplatesPath.Size = New Size(700, 20)
        txtTemplatesPath.ReadOnly = True
        grpConfig.Controls.Add(txtTemplatesPath)

        Dim lblOutput As New Label()
        lblOutput.Text = "Dossier Output:"
        lblOutput.Location = New Point(10, 50)
        lblOutput.Size = New Size(120, 20)
        grpConfig.Controls.Add(lblOutput)

        txtOutputPath = New TextBox()
        txtOutputPath.Location = New Point(140, 48)
        txtOutputPath.Size = New Size(700, 20)
        txtOutputPath.ReadOnly = True
        grpConfig.Controls.Add(txtOutputPath)

        Dim lblPara As New Label()
        lblPara.Text = "Template paragraphes:"
        lblPara.Location = New Point(10, 75)
        lblPara.Size = New Size(120, 20)
        grpConfig.Controls.Add(lblPara)

        txtParagraphTemplate = New TextBox()
        txtParagraphTemplate.Location = New Point(140, 73)
        txtParagraphTemplate.Size = New Size(700, 20)
        txtParagraphTemplate.ReadOnly = True
        grpConfig.Controls.Add(txtParagraphTemplate)

        Me.Controls.Add(grpConfig)

        ' GroupBox - Génération
        Dim grpGeneration As New GroupBox()
        grpGeneration.Text = "3. Génération des documents"
        grpGeneration.Location = New Point(10, 230)
        grpGeneration.Size = New Size(860, 80)

        btnGenerateBDO = New Button()
        btnGenerateBDO.Text = "Générer BDO"
        btnGenerateBDO.Location = New Point(10, 25)
        btnGenerateBDO.Size = New Size(180, 40)
        grpGeneration.Controls.Add(btnGenerateBDO)

        btnGenerateContracts = New Button()
        btnGenerateContracts.Text = "Générer Contrats"
        btnGenerateContracts.Location = New Point(200, 25)
        btnGenerateContracts.Size = New Size(180, 40)
        grpGeneration.Controls.Add(btnGenerateContracts)

        btnGenerateAll = New Button()
        btnGenerateAll.Text = "TOUT GÉNÉRER"
        btnGenerateAll.Location = New Point(390, 25)
        btnGenerateAll.Size = New Size(180, 40)
        btnGenerateAll.Font = New Font(btnGenerateAll.Font, FontStyle.Bold)
        grpGeneration.Controls.Add(btnGenerateAll)

        btnOpenOutput = New Button()
        btnOpenOutput.Text = "Ouvrir dossier Output"
        btnOpenOutput.Location = New Point(660, 25)
        btnOpenOutput.Size = New Size(180, 40)
        grpGeneration.Controls.Add(btnOpenOutput)

        Me.Controls.Add(grpGeneration)

        ' GroupBox - Logs
        Dim grpLogs As New GroupBox()
        grpLogs.Text = "4. Logs de génération"
        grpLogs.Location = New Point(10, 320)
        grpLogs.Size = New Size(860, 320)

        txtLog = New TextBox()
        txtLog.Multiline = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.Location = New Point(10, 25)
        txtLog.Size = New Size(830, 250)
        txtLog.Font = New Font("Consolas", 9)
        txtLog.ReadOnly = True
        grpLogs.Controls.Add(txtLog)

        btnClearLog = New Button()
        btnClearLog.Text = "Effacer les logs"
        btnClearLog.Location = New Point(10, 280)
        btnClearLog.Size = New Size(150, 30)
        grpLogs.Controls.Add(btnClearLog)

        Me.Controls.Add(grpLogs)

        ' Initialiser les logs
        btnClearLog_Click(Nothing, Nothing)
    End Sub

    ' Contrôles du formulaire
    Private WithEvents txtJsonPath As TextBox
    Private WithEvents btnSelectJson As Button
    Private WithEvents lblTitre As Label
    Private WithEvents lblInterprete As Label
    Private WithEvents lblAyantsDroit As Label
    Private WithEvents txtTemplatesPath As TextBox
    Private WithEvents txtOutputPath As TextBox
    Private WithEvents txtParagraphTemplate As TextBox
    Private WithEvents btnGenerateBDO As Button
    Private WithEvents btnGenerateContracts As Button
    Private WithEvents btnGenerateAll As Button
    Private WithEvents btnOpenOutput As Button
    Private WithEvents txtLog As TextBox
    Private WithEvents btnClearLog As Button
End Class
