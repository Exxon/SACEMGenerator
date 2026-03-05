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
        
        ' Réinitialiser les indicateurs
        ResetIndicators()
    End Sub

    ''' <summary>
    ''' Réinitialise tous les indicateurs à leurs valeurs par défaut
    ''' </summary>
    Private Sub ResetIndicators()
        lblTitre.Text = "Titre: -"
        lblInterprete.Text = "Interprète: -"
        lblAyantsDroit.Text = "Ayants droit: 0"
        lblLettrages.Text = "Lettrages: 0"
        lblAuteurs.Text = "A: 0"
        lblCompositeurs.Text = "C: 0"
        lblEditeurs.Text = "E: 0"
        
        ' Indicateurs NON-SACEM
        lblNonSACEM.Text = "NON-SACEM: 0"
        lblNonSACEM.ForeColor = Color.Black
        lblPartsInedites.Text = "Parts inédites: 0"
        lblPartsInedites.ForeColor = Color.Black
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

            ' Calculer et afficher les statistiques détaillées
            UpdateDetailedStatistics()

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
            
            ResetIndicators()
        End Try
    End Sub

    ''' <summary>
    ''' Calcule et affiche les statistiques détaillées des ayants droit
    ''' </summary>
    Private Sub UpdateDetailedStatistics()
        If _currentData Is Nothing OrElse _currentData.AyantsDroit Is Nothing Then
            ResetIndicators()
            Return
        End If
        
        ' Compteurs
        Dim countA As Integer = 0      ' Auteurs
        Dim countC As Integer = 0      ' Compositeurs
        Dim countE As Integer = 0      ' Éditeurs
        Dim countNonSACEM As Integer = 0
        Dim countPartsInedites As Integer = 0
        
        ' Liste des lettrages uniques
        Dim lettrages As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        
        ' Dictionnaire pour compter les ayants droit par lettrage (pour détecter parts inédites)
        Dim ayantsDroitParLettrage As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        
        ' Parcourir tous les ayants droit
        For Each ayant In _currentData.AyantsDroit
            Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
            Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
            Dim societe As String = If(ayant.Identite.SocieteGestion, "SACEM").Trim().ToUpper()
            
            ' Compter par rôle
            Select Case role
                Case "A", "AD"
                    countA += 1
                Case "C", "AR"
                    countC += 1
                Case "AC"
                    countA += 1
                    countC += 1
                Case "E"
                    countE += 1
            End Select
            
            ' Collecter les lettrages
            If Not String.IsNullOrEmpty(lettrage) Then
                lettrages.Add(lettrage)
                
                ' Grouper par lettrage pour détecter les parts inédites
                If Not ayantsDroitParLettrage.ContainsKey(lettrage) Then
                    ayantsDroitParLettrage(lettrage) = New List(Of String)
                End If
                ayantsDroitParLettrage(lettrage).Add(role)
            End If
            
            ' Compter NON-SACEM
            If societe <> "SACEM" AndAlso Not String.IsNullOrEmpty(societe) Then
                countNonSACEM += 1
            End If
        Next
        
        ' Détecter les parts inédites (AC seul dans son lettrage = pas d'éditeur)
        For Each kvp In ayantsDroitParLettrage
            Dim roles As List(Of String) = kvp.Value
            Dim hasAC As Boolean = roles.Any(Function(r) r = "A" OrElse r = "C" OrElse r = "AC" OrElse r = "AR" OrElse r = "AD")
            Dim hasE As Boolean = roles.Any(Function(r) r = "E")
            
            ' Si un lettrage a des AC mais pas d'éditeur = part inédite
            If hasAC AndAlso Not hasE Then
                ' Compter le nombre d'AC dans ce lettrage
                countPartsInedites += roles.Where(Function(r) r = "A" OrElse r = "C" OrElse r = "AC" OrElse r = "AR" OrElse r = "AD").Count()
            End If
        Next
        
        ' Afficher les compteurs
        lblLettrages.Text = $"Lettrages: {lettrages.Count}"
        lblAuteurs.Text = $"A: {countA}"
        lblCompositeurs.Text = $"C: {countC}"
        lblEditeurs.Text = $"E: {countE}"
        
        ' Afficher NON-SACEM avec indicateur visuel
        lblNonSACEM.Text = $"NON-SACEM: {countNonSACEM}"
        If countNonSACEM > 0 Then
            lblNonSACEM.ForeColor = Color.OrangeRed
            lblNonSACEM.Font = New Font(lblNonSACEM.Font, FontStyle.Bold)
        Else
            lblNonSACEM.ForeColor = Color.Green
            lblNonSACEM.Font = New Font(lblNonSACEM.Font, FontStyle.Regular)
        End If
        
        ' Afficher Parts inédites avec indicateur visuel
        lblPartsInedites.Text = $"Parts inédites: {countPartsInedites}"
        If countPartsInedites > 0 Then
            lblPartsInedites.ForeColor = Color.OrangeRed
            lblPartsInedites.Font = New Font(lblPartsInedites.Font, FontStyle.Bold)
        Else
            lblPartsInedites.ForeColor = Color.Green
            lblPartsInedites.Font = New Font(lblPartsInedites.Font, FontStyle.Regular)
        End If
        
        ' Log des statistiques
        txtLog.AppendText($"{vbCrLf}=== STATISTIQUES ==={vbCrLf}")
        txtLog.AppendText($"  Lettrages: {lettrages.Count} ({String.Join(", ", lettrages.OrderBy(Function(l) l))}){vbCrLf}")
        txtLog.AppendText($"  Auteurs (A/AD): {countA}{vbCrLf}")
        txtLog.AppendText($"  Compositeurs (C/AR): {countC}{vbCrLf}")
        txtLog.AppendText($"  Éditeurs (E): {countE}{vbCrLf}")
        txtLog.AppendText($"  NON-SACEM: {countNonSACEM}{vbCrLf}")
        txtLog.AppendText($"  Parts inédites: {countPartsInedites}{vbCrLf}")
        
        ' Afficher l'alerte NON-SACEM si nécessaire
        If countNonSACEM > 0 Then
            txtLog.AppendText($"  ⚠ ŒUVRE MIXTE détectée (membres NON-SACEM présents){vbCrLf}")
            
            ' Afficher la MessageBox d'alerte
            ShowNonSACEMAlert()
        End If
    End Sub

    ''' <summary>
    ''' Affiche une alerte détaillée pour les ayants droit NON-SACEM
    ''' </summary>
    Private Sub ShowNonSACEMAlert()
        If _currentData Is Nothing OrElse _currentData.AyantsDroit Is Nothing Then Return
        
        Dim alertMessage As New System.Text.StringBuilder()
        alertMessage.AppendLine("⚠ ŒUVRE MIXTE DÉTECTÉE")
        alertMessage.AppendLine()
        alertMessage.AppendLine("Les ayants droit suivants ne sont pas membres de la SACEM :")
        alertMessage.AppendLine()
        
        ' Collecter les NON-SACEM par type (AC vs E)
        Dim acNonSACEM As New List(Of String)
        Dim eNonSACEM As New List(Of String)
        
        For Each ayant In _currentData.AyantsDroit
            Dim societe As String = If(ayant.Identite.SocieteGestion, "").Trim().ToUpper()
            
            ' Si vide ou SACEM, on ignore
            If String.IsNullOrEmpty(societe) OrElse societe = "SACEM" Then Continue For
            
            Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
            Dim isAC As Boolean = (role = "A" OrElse role = "C" OrElse role = "AC" OrElse role = "AR" OrElse role = "AD")
            Dim isE As Boolean = (role = "E")
            
            ' Nom d'affichage
            Dim displayName As String
            If ayant.Identite.Type = "Physique" Then
                displayName = $"{ayant.Identite.Prenom} {ayant.Identite.Nom}".Trim()
                If String.IsNullOrEmpty(displayName) Then displayName = ayant.Identite.Designation
            Else
                displayName = ayant.Identite.Designation
            End If
            
            Dim info As String = $"• {displayName} ({ayant.Identite.SocieteGestion})"
            
            If isAC Then
                acNonSACEM.Add(info)
            ElseIf isE Then
                eNonSACEM.Add(info)
            End If
        Next
        
        ' Afficher les AC NON-SACEM
        If acNonSACEM.Count > 0 Then
            alertMessage.AppendLine("AUTEURS/COMPOSITEURS :")
            For Each ac In acNonSACEM
                alertMessage.AppendLine(ac)
            Next
            alertMessage.AppendLine()
            alertMessage.AppendLine("→ Ne signeront PAS : CCEOM, CCDAA")
            alertMessage.AppendLine("→ Signeront UNIQUEMENT : Split Sheet (Lettre de Répartition)")
            alertMessage.AppendLine()
        End If
        
        ' Afficher les E NON-SACEM
        If eNonSACEM.Count > 0 Then
            alertMessage.AppendLine("ÉDITEURS :")
            For Each e In eNonSACEM
                alertMessage.AppendLine(e)
            Next
            alertMessage.AppendLine()
            alertMessage.AppendLine("→ Ne signeront PAS : COED")
            alertMessage.AppendLine("→ Signeront UNIQUEMENT : Split Sheet (Lettre de Répartition)")
            alertMessage.AppendLine()
        End If
        
        alertMessage.AppendLine("────────────────────────────────")
        alertMessage.AppendLine("Les mentions NON-SACEM seront ajoutées automatiquement")
        alertMessage.AppendLine("dans les contrats (Articles 11 et 16 du CCEOM, Article 3 du COED).")
        
        ' Afficher la MessageBox
        MessageBox.Show(alertMessage.ToString(), 
                       "Alerte : Membres NON-SACEM détectés", 
                       MessageBoxButtons.OK, 
                       MessageBoxIcon.Warning)
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

            ' Vérifier le template BDO PDF officiel SACEM
            Dim bdoTemplatePath As String = Path.Combine(_templatesDirectory, "Bdo711.pdf")
            If Not File.Exists(bdoTemplatePath) Then
                MessageBox.Show($"Template PDF BDO introuvable:{vbCrLf}{bdoTemplatePath}", 
                               "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Nom du fichier de sortie
            Dim outputFileName As String = $"BDO_{_currentData.Titre}_{_currentData.Interprete}.pdf"
            outputFileName = CleanFileName(outputFileName)
            Dim outputPath As String = Path.Combine(_outputDirectory, outputFileName)

            ' Générer le BDO avec le nouveau générateur PDF
            Dim generator As New BDOPdfGenerator(_currentData)
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
                MessageBox.Show($"Templates manquants:{vbCrLf}{String.Join(vbCrLf, missingTemplates)}", 
                               "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Générer chaque contrat
            For Each kvp In templates
                Dim contractType As String = kvp.Key
                Dim templatePath As String = kvp.Value
                Dim outputFileName As String = $"{contractType}_{_currentData.Titre}_{_currentData.Interprete}.docx"
                outputFileName = CleanFileName(outputFileName)
                Dim outputPath As String = Path.Combine(_outputDirectory, outputFileName)

                txtLog.AppendText($"Génération {contractType}...{vbCrLf}")

                ' Créer le générateur de contrat avec les données et le reader de paragraphes
                Dim contractGenerator As New ContractGenerator(_currentData, _paragraphReader)
                Dim success As Boolean = contractGenerator.Generate(templatePath, outputPath, contractType)

                ' Afficher les logs du générateur
                For Each logEntry In contractGenerator.GenerationLog
                    txtLog.AppendText($"  {logEntry}{vbCrLf}")
                Next

                If success Then
                    txtLog.AppendText($"✓ {contractType} généré: {outputFileName}{vbCrLf}")
                Else
                    txtLog.AppendText($"✗ Échec génération {contractType}{vbCrLf}")
                End If
            Next

            MessageBox.Show("Génération des contrats terminée. Consultez les logs.", 
                           "Terminé", MessageBoxButtons.OK, MessageBoxIcon.Information)

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
            txtLog.AppendText(vbCrLf & New String("="c, 50) & vbCrLf)
            txtLog.AppendText("=== GÉNÉRATION COMPLÈTE ===" & vbCrLf)
            txtLog.AppendText(New String("="c, 50) & vbCrLf)

            ' Générer BDO
            btnGenerateBDO_Click(sender, e)

            ' Générer Contrats
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
    Private Sub btnEditJson_Click(sender As Object, e As EventArgs) Handles btnEditJson.Click
        Dim editorPath As String = If(_currentJsonPath, "")
        Using editor As New JsonEditorForm(editorPath)
            editor.ShowDialog()
            If Not String.IsNullOrEmpty(editor.SavedJsonPath) AndAlso
               File.Exists(editor.SavedJsonPath) Then
                _currentJsonPath = editor.SavedJsonPath
                txtJsonPath.Text = _currentJsonPath
                LoadAndValidateJson()
            End If
        End Using
    End Sub

    Private Sub btnClearLog_Click(sender As Object, e As EventArgs) Handles btnClearLog.Click
        txtLog.Clear()
        txtLog.AppendText("=== SACEM GENERATOR - VB.NET ===" & vbCrLf)
        txtLog.AppendText($"Version 1.1 - {DateTime.Now:yyyy-MM-dd}" & vbCrLf)
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
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = "SACEM Generator - VB.NET"
        Me.Size = New Size(900, 780)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' GroupBox - Sélection JSON
        Dim grpJson As New GroupBox()
        grpJson.Text = "1. Sélection du fichier JSON"
        grpJson.Location = New Point(10, 10)
        grpJson.Size = New Size(860, 155)

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

        btnEditJson = New Button()
        btnEditJson.Text = "Créer / Éditer JSON"
        btnEditJson.Location = New Point(720, 55)
        btnEditJson.Size = New Size(120, 25)
        btnEditJson.BackColor = Color.FromArgb(16, 124, 16)
        btnEditJson.ForeColor = Color.White
        btnEditJson.FlatStyle = FlatStyle.Flat
        grpJson.Controls.Add(btnEditJson)

        ' Ligne 1 : Titre et Interprète
        lblTitre = New Label()
        lblTitre.Location = New Point(10, 55)
        lblTitre.Size = New Size(400, 20)
        lblTitre.Text = "Titre: -"
        grpJson.Controls.Add(lblTitre)

        lblInterprete = New Label()
        lblInterprete.Location = New Point(420, 55)
        lblInterprete.Size = New Size(400, 20)
        lblInterprete.Text = "Interprète: -"
        grpJson.Controls.Add(lblInterprete)

        ' Ligne 2 : Ayants droit, Lettrages, A, C, E
        lblAyantsDroit = New Label()
        lblAyantsDroit.Location = New Point(10, 80)
        lblAyantsDroit.Size = New Size(120, 20)
        lblAyantsDroit.Text = "Ayants droit: 0"
        grpJson.Controls.Add(lblAyantsDroit)

        lblLettrages = New Label()
        lblLettrages.Location = New Point(140, 80)
        lblLettrages.Size = New Size(100, 20)
        lblLettrages.Text = "Lettrages: 0"
        grpJson.Controls.Add(lblLettrages)

        lblAuteurs = New Label()
        lblAuteurs.Location = New Point(250, 80)
        lblAuteurs.Size = New Size(60, 20)
        lblAuteurs.Text = "A: 0"
        lblAuteurs.ForeColor = Color.DarkBlue
        grpJson.Controls.Add(lblAuteurs)

        lblCompositeurs = New Label()
        lblCompositeurs.Location = New Point(320, 80)
        lblCompositeurs.Size = New Size(60, 20)
        lblCompositeurs.Text = "C: 0"
        lblCompositeurs.ForeColor = Color.DarkBlue
        grpJson.Controls.Add(lblCompositeurs)

        lblEditeurs = New Label()
        lblEditeurs.Location = New Point(390, 80)
        lblEditeurs.Size = New Size(60, 20)
        lblEditeurs.Text = "E: 0"
        lblEditeurs.ForeColor = Color.DarkBlue
        grpJson.Controls.Add(lblEditeurs)

        ' Ligne 3 : NON-SACEM et Parts inédites (avec indicateurs visuels)
        lblNonSACEM = New Label()
        lblNonSACEM.Location = New Point(10, 105)
        lblNonSACEM.Size = New Size(150, 20)
        lblNonSACEM.Text = "NON-SACEM: 0"
        grpJson.Controls.Add(lblNonSACEM)

        lblPartsInedites = New Label()
        lblPartsInedites.Location = New Point(170, 105)
        lblPartsInedites.Size = New Size(150, 20)
        lblPartsInedites.Text = "Parts inédites: 0"
        grpJson.Controls.Add(lblPartsInedites)

        Me.Controls.Add(grpJson)

        ' GroupBox - Configuration
        Dim grpConfig As New GroupBox()
        grpConfig.Text = "2. Configuration des chemins"
        grpConfig.Location = New Point(10, 175)
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
        grpGeneration.Location = New Point(10, 285)
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
        grpLogs.Location = New Point(10, 375)
        grpLogs.Size = New Size(860, 350)

        txtLog = New TextBox()
        txtLog.Multiline = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.Location = New Point(10, 25)
        txtLog.Size = New Size(830, 280)
        txtLog.Font = New Font("Consolas", 9)
        txtLog.ReadOnly = True
        grpLogs.Controls.Add(txtLog)

        btnClearLog = New Button()
        btnClearLog.Text = "Effacer les logs"
        btnClearLog.Location = New Point(10, 310)
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
    Private WithEvents lblLettrages As Label
    Private WithEvents lblAuteurs As Label
    Private WithEvents lblCompositeurs As Label
    Private WithEvents lblEditeurs As Label
    Private WithEvents lblNonSACEM As Label
    Private WithEvents lblPartsInedites As Label
    Private WithEvents txtTemplatesPath As TextBox
    Private WithEvents txtOutputPath As TextBox
    Private WithEvents txtParagraphTemplate As TextBox
    Private WithEvents btnGenerateBDO As Button
    Private WithEvents btnGenerateContracts As Button
    Private WithEvents btnGenerateAll As Button
    Private WithEvents btnOpenOutput As Button
    Private WithEvents txtLog As TextBox
    Private WithEvents btnClearLog As Button
    Private WithEvents btnEditJson As Button
End Class
