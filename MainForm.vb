Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq
Imports OfficeOpenXml

''' <summary>
''' Formulaire principal unifié SACEM Generator.
''' Fusionne l'éditeur JSON (œuvre + ayants droit) et la génération de documents.
''' Layout : Gauche = Œuvre | Milieu = Ayants droit | Droite = Stats + Config + Génération + Logs
''' </summary>
Public Class MainForm
    Inherits Form

    ' ─────────────────────────────────────────────────────────────
    ' DONNÉES — GÉNÉRATION (ex-MainForm)
    ' ─────────────────────────────────────────────────────────────
    Private _currentJsonPath As String
    Private _currentData As SACEMData
    Private _paragraphReader As ParagraphTemplateReader
    Private _templatesDirectory As String = "Templates"
    Private _outputDirectory As String = "Output"
    Private _paragraphTemplatePath As String = "Templates\template_paragrahs.docx"

    ' ─────────────────────────────────────────────────────────────
    ' DONNÉES — ÉDITEUR (ex-JsonEditorForm)
    ' ─────────────────────────────────────────────────────────────
    Private DtPersonPhy As DataTable
    Private DtPersonMor As DataTable
    Private DtDepotCreateur As DataTable
    Private _dragRowIndex As Integer = -1
    Private _totalAvantPH As Double = 0.0
    Private _watermark As Boolean = True

    ' ─────────────────────────────────────────────────────────────
    ' CONTRÔLES — PANNEAU GAUCHE (Œuvre)
    ' ─────────────────────────────────────────────────────────────
    Private WithEvents txtTitre As TextBox
    Private WithEvents txtSousTitre As TextBox
    Private WithEvents txtInterprete As TextBox
    Private WithEvents txtDuree As TextBox
    Private WithEvents cbGenre As ComboBox
    Private WithEvents dtDate As DateTimePicker
    Private WithEvents cbLieu As ComboBox
    Private WithEvents cbTerritoire As ComboBox
    Private WithEvents cbArrangement As ComboBox
    Private WithEvents txtISWC As TextBox
    Private WithEvents cbDeclaration As ComboBox
    Private WithEvents cbFormat As ComboBox
    Private WithEvents txtFaita As TextBox
    Private WithEvents dtFaitle As DateTimePicker
    Private WithEvents cbInegalitaire As CheckBox
    Private WithEvents cbPersonnes As ComboBox

    ' ─────────────────────────────────────────────────────────────
    ' CONTRÔLES — PANNEAU MILIEU (Ayants droit)
    ' ─────────────────────────────────────────────────────────────
    Private WithEvents txtRecherche As TextBox
    Private WithEvents lstResultats As ListBox
    Private WithEvents btnAjouter As Button
    Private WithEvents btnCalculer As Button
    Private WithEvents btnGererFiches As Button
    Private WithEvents dgv As DataGridView
    Private WithEvents cmsGrille As ContextMenuStrip
    Private WithEvents mnuSupprimer As ToolStripMenuItem
    Private WithEvents lblStatut As Label

    ' ─────────────────────────────────────────────────────────────
    ' CONTRÔLES — PANNEAU DROIT (Stats + Config + Génération + Logs)
    ' ─────────────────────────────────────────────────────────────
    Private WithEvents txtJsonPath As TextBox
    Private WithEvents btnSelectJson As Button
    Private WithEvents btnChargerJson As Button
    Private WithEvents btnSauvegarder As Button
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
    Private splitMain As SplitContainer
    Private splitRight As SplitContainer

    ' ─────────────────────────────────────────────────────────────
    ' CONSTRUCTEUR
    ' ─────────────────────────────────────────────────────────────
    Public Sub New()
        InitializeComponent()
        InitializePaths()
        PersonnesForm.InitialiserBDD()
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' INITIALISATION DES COMPOSANTS
    ' ─────────────────────────────────────────────────────────────
    Private Sub InitializeComponent()
        Me.Text = "SACEM Generator"
        Me.WindowState = FormWindowState.Maximized
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(1100, 700)
        Me.BackColor = Color.FromArgb(240, 242, 245)

        ' ── SplitContainer gauche/reste ─────────────────────────
        splitMain = New SplitContainer()
        splitMain.Dock = DockStyle.Fill
        splitMain.Orientation = Orientation.Vertical
        splitMain.SplitterWidth = 5
        splitMain.BackColor = Color.FromArgb(210, 218, 230)
        Me.Controls.Add(splitMain)

        ' ── SplitContainer milieu/droite ────────────────────────
        splitRight = New SplitContainer()
        splitRight.Dock = DockStyle.Fill
        splitRight.Orientation = Orientation.Vertical
        splitRight.SplitterWidth = 5
        splitRight.BackColor = Color.FromArgb(210, 218, 230)
        splitMain.Panel2.Controls.Add(splitRight)

        ' ════════════════════════════════════════════════════════
        ' PANNEAU GAUCHE — Œuvre
        ' ════════════════════════════════════════════════════════
        splitMain.Panel1.BackColor = Color.FromArgb(245, 247, 250)
        splitMain.Panel1.AutoScroll = True

        Dim pnlOeuvre As New Panel()
        pnlOeuvre.Location = New Point(8, 8)
        pnlOeuvre.Width = 296
        pnlOeuvre.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        pnlOeuvre.BackColor = Color.White
        splitMain.Panel1.Controls.Add(pnlOeuvre)

        Dim lblOH As New Label()
        lblOH.Text = "OEUVRE"
        lblOH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblOH.ForeColor = Color.FromArgb(30, 80, 160)
        lblOH.Location = New Point(12, 10)
        lblOH.Size = New Size(270, 18)
        pnlOeuvre.Controls.Add(lblOH)

        Dim sepO As New Panel()
        sepO.Location = New Point(12, 30)
        sepO.Size = New Size(272, 1)
        sepO.BackColor = Color.FromArgb(200, 215, 235)
        pnlOeuvre.Controls.Add(sepO)

        Dim yO As Integer = 38
        Dim fw As Integer = 272

        pnlOeuvre.Controls.Add(MkCaption("Titre", 12, yO))
        txtTitre = MkTb(12, yO + 15, fw)
        pnlOeuvre.Controls.Add(txtTitre)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Sous-titre", 12, yO))
        txtSousTitre = MkTb(12, yO + 15, fw)
        pnlOeuvre.Controls.Add(txtSousTitre)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Interprete", 12, yO))
        txtInterprete = MkTb(12, yO + 15, fw)
        pnlOeuvre.Controls.Add(txtInterprete)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Genre", 12, yO))
        pnlOeuvre.Controls.Add(MkCaption("Duree", 196, yO))
        cbGenre = New ComboBox()
        cbGenre.Location = New Point(12, yO + 15)
        cbGenre.Size = New Size(176, 23)
        cbGenre.Items.AddRange(New Object() {"Chant", "Jazz", "Techno", "Musique de film",
            "Musique symphonique", "Instrumental", "Musique illustrative",
            "Billet d'humeur", "Texte", "Texte de presentation",
            "Texte de sketch", "Poeme", "Chronique"})
        cbGenre.Text = "Chant"
        pnlOeuvre.Controls.Add(cbGenre)
        txtDuree = MkTb(192, yO + 15, 92)
        pnlOeuvre.Controls.Add(txtDuree)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("ISWC", 12, yO))
        txtISWC = MkTb(12, yO + 15, fw)
        pnlOeuvre.Controls.Add(txtISWC)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Date d'exploitation", 12, yO))
        dtDate = New DateTimePicker()
        dtDate.Location = New Point(12, yO + 15)
        dtDate.Size = New Size(fw, 23)
        dtDate.Format = DateTimePickerFormat.Custom
        dtDate.CustomFormat = "dd/MM/yyyy"
        pnlOeuvre.Controls.Add(dtDate)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Lieu", 12, yO))
        cbLieu = New ComboBox()
        cbLieu.Location = New Point(12, yO + 15)
        cbLieu.Size = New Size(fw, 23)
        cbLieu.Items.Add("Deezer / Spotify / Amazon Music / Youtube/ Apple Music")
        pnlOeuvre.Controls.Add(cbLieu)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Territoire", 12, yO))
        cbTerritoire = New ComboBox()
        cbTerritoire.Location = New Point(12, yO + 15)
        cbTerritoire.Size = New Size(fw, 23)
        cbTerritoire.Items.Add("Monde")
        pnlOeuvre.Controls.Add(cbTerritoire)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Arrangement", 12, yO))
        cbArrangement = New ComboBox()
        cbArrangement.Location = New Point(12, yO + 15)
        cbArrangement.Size = New Size(fw, 23)
        cbArrangement.Items.Add("Toutes")
        pnlOeuvre.Controls.Add(cbArrangement)
        yO += 44

        cbInegalitaire = New CheckBox()
        cbInegalitaire.Text = "Repartition inegalitaire"
        cbInegalitaire.Location = New Point(12, yO + 4)
        cbInegalitaire.Size = New Size(220, 22)
        cbInegalitaire.Font = New Font("Segoe UI", 9)
        pnlOeuvre.Controls.Add(cbInegalitaire)
        yO += 36

        pnlOeuvre.Controls.Add(MkCaption("Declaration", 12, yO))
        cbDeclaration = New ComboBox()
        cbDeclaration.Location = New Point(12, yO + 15)
        cbDeclaration.Size = New Size(fw, 23)
        cbDeclaration.DropDownStyle = ComboBoxStyle.DropDown
        pnlOeuvre.Controls.Add(cbDeclaration)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Format", 12, yO))
        cbFormat = New ComboBox()
        cbFormat.Location = New Point(12, yO + 15)
        cbFormat.Size = New Size(fw, 23)
        cbFormat.DropDownStyle = ComboBoxStyle.DropDown
        pnlOeuvre.Controls.Add(cbFormat)
        yO += 44

        pnlOeuvre.Controls.Add(MkCaption("Fait a", 12, yO))
        pnlOeuvre.Controls.Add(MkCaption("Fait le", 152, yO))
        txtFaita = MkTb(12, yO + 15, 132)
        pnlOeuvre.Controls.Add(txtFaita)
        dtFaitle = New DateTimePicker()
        dtFaitle.Location = New Point(152, yO + 15)
        dtFaitle.Size = New Size(132, 23)
        dtFaitle.Format = DateTimePickerFormat.Custom
        dtFaitle.CustomFormat = "dd/MM/yyyy"
        pnlOeuvre.Controls.Add(dtFaitle)
        yO += 52

        pnlOeuvre.Height = yO + 20

        ' cbPersonnes invisible (compatibilité PopulateComboBox)
        cbPersonnes = New ComboBox()
        cbPersonnes.Visible = False
        splitMain.Panel1.Controls.Add(cbPersonnes)

        ' ════════════════════════════════════════════════════════
        ' PANNEAU MILIEU — Ayants droit
        ' ════════════════════════════════════════════════════════
        splitRight.Panel1.BackColor = Color.FromArgb(245, 247, 250)

        Dim pnlMidWrap As New Panel()
        pnlMidWrap.Dock = DockStyle.Fill
        pnlMidWrap.Padding = New Padding(8)
        pnlMidWrap.BackColor = Color.FromArgb(245, 247, 250)
        splitRight.Panel1.Controls.Add(pnlMidWrap)

        Dim pnlAD As New Panel()
        pnlAD.Dock = DockStyle.Fill
        pnlAD.BackColor = Color.White
        pnlMidWrap.Controls.Add(pnlAD)

        Dim lblADH As New Label()
        lblADH.Text = "AYANTS DROIT"
        lblADH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblADH.ForeColor = Color.FromArgb(30, 80, 160)
        lblADH.Location = New Point(12, 10)
        lblADH.Size = New Size(300, 18)
        pnlAD.Controls.Add(lblADH)

        Dim sepAD As New Panel()
        sepAD.Location = New Point(12, 30)
        sepAD.Size = New Size(900, 1)
        sepAD.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        sepAD.BackColor = Color.FromArgb(200, 215, 235)
        pnlAD.Controls.Add(sepAD)

        txtRecherche = New TextBox()
        txtRecherche.Location = New Point(12, 40)
        txtRecherche.Size = New Size(240, 26)
        txtRecherche.ForeColor = Color.Gray
        txtRecherche.Text = "Nom, prenom, pseudonyme..."
        txtRecherche.Font = New Font("Segoe UI", 9)
        pnlAD.Controls.Add(txtRecherche)

        btnAjouter = MkBtn("+ AJOUTER", 260, 40, 100, Color.FromArgb(34, 139, 34))
        pnlAD.Controls.Add(btnAjouter)

        btnCalculer = MkBtn("CALCULER %", 368, 40, 110, Color.FromArgb(180, 100, 20))
        pnlAD.Controls.Add(btnCalculer)

        btnGererFiches = MkBtn("Gerer les fiches BDD", 486, 40, 160, Color.FromArgb(107, 60, 157))
        pnlAD.Controls.Add(btnGererFiches)

        lstResultats = New ListBox()
        lstResultats.Location = New Point(12, 70)
        lstResultats.Size = New Size(350, 100)
        lstResultats.Visible = False
        lstResultats.Font = New Font("Segoe UI", 9)
        lstResultats.BorderStyle = BorderStyle.FixedSingle
        pnlAD.Controls.Add(lstResultats)

        cmsGrille = New ContextMenuStrip()
        mnuSupprimer = New ToolStripMenuItem("Supprimer la ligne")
        cmsGrille.Items.Add(mnuSupprimer)

        dgv = New DataGridView()
        dgv.Location = New Point(12, 74)
        dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        dgv.Size = New Size(600, 700)
        dgv.ContextMenuStrip = cmsGrille
        dgv.AllowUserToAddRows = False
        dgv.ReadOnly = False
        dgv.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
        dgv.RowTemplate.Height = 26
        dgv.ColumnHeadersHeight = 28
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 80, 160)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 242, 255)
        dgv.BorderStyle = BorderStyle.None
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgv.GridColor = Color.FromArgb(210, 220, 235)
        dgv.AllowDrop = True
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.BackgroundColor = Color.White
        pnlAD.Controls.Add(dgv)

        lblStatut = New Label()
        lblStatut.Dock = DockStyle.Bottom
        lblStatut.Height = 22
        lblStatut.Text = "Pret."
        lblStatut.Font = New Font("Segoe UI", 8)
        lblStatut.ForeColor = Color.FromArgb(80, 100, 130)
        lblStatut.TextAlign = ContentAlignment.MiddleLeft
        lblStatut.Padding = New Padding(4, 0, 0, 0)
        pnlAD.Controls.Add(lblStatut)

        ' ════════════════════════════════════════════════════════
        ' PANNEAU DROIT — Stats + Config + JSON + Génération + Logs
        ' ════════════════════════════════════════════════════════
        splitRight.Panel2.BackColor = Color.FromArgb(245, 247, 250)
        splitRight.Panel2.AutoScroll = True

        Dim pnlRight As New Panel()
        pnlRight.Location = New Point(8, 8)
        pnlRight.Width = 420
        pnlRight.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        pnlRight.BackColor = Color.White
        splitRight.Panel2.Controls.Add(pnlRight)

        Dim yR As Integer = 10
        Dim rw As Integer = 396

        ' ── Section : JSON ───────────────────────────────────────
        Dim lblJH As New Label()
        lblJH.Text = "FICHIER JSON"
        lblJH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblJH.ForeColor = Color.FromArgb(30, 80, 160)
        lblJH.Location = New Point(12, yR)
        lblJH.Size = New Size(rw, 18)
        pnlRight.Controls.Add(lblJH)
        yR += 22

        Dim sepJ As New Panel()
        sepJ.Location = New Point(12, yR)
        sepJ.Size = New Size(rw, 1)
        sepJ.BackColor = Color.FromArgb(200, 215, 235)
        pnlRight.Controls.Add(sepJ)
        yR += 8

        txtJsonPath = New TextBox()
        txtJsonPath.Location = New Point(12, yR)
        txtJsonPath.Size = New Size(rw, 22)
        txtJsonPath.ReadOnly = True
        txtJsonPath.Font = New Font("Segoe UI", 8)
        txtJsonPath.ForeColor = Color.FromArgb(80, 80, 80)
        pnlRight.Controls.Add(txtJsonPath)
        yR += 30

        btnChargerJson = MkBtn("Charger JSON", 12, yR, 162, Color.FromArgb(70, 110, 170))
        pnlRight.Controls.Add(btnChargerJson)

        btnSauvegarder = MkBtn("Sauvegarder JSON", 182, yR, 170, Color.FromArgb(0, 120, 212))
        btnSauvegarder.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        pnlRight.Controls.Add(btnSauvegarder)
        yR += 36

        btnSelectJson = MkBtn("Parcourir JSON existant...", 12, yR, rw, Color.FromArgb(100, 120, 150))
        pnlRight.Controls.Add(btnSelectJson)
        yR += 42

        ' ── Section : Statistiques ───────────────────────────────
        Dim sepSt As New Panel()
        sepSt.Location = New Point(12, yR)
        sepSt.Size = New Size(rw, 1)
        sepSt.BackColor = Color.FromArgb(200, 215, 235)
        pnlRight.Controls.Add(sepSt)
        yR += 8

        Dim lblStH As New Label()
        lblStH.Text = "STATISTIQUES"
        lblStH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblStH.ForeColor = Color.FromArgb(30, 80, 160)
        lblStH.Location = New Point(12, yR)
        lblStH.Size = New Size(rw, 18)
        pnlRight.Controls.Add(lblStH)
        yR += 24

        lblTitre = New Label()
        lblTitre.Text = "Titre: -"
        lblTitre.Location = New Point(12, yR)
        lblTitre.Size = New Size(rw, 18)
        lblTitre.Font = New Font("Segoe UI", 8)
        pnlRight.Controls.Add(lblTitre)
        yR += 20

        lblInterprete = New Label()
        lblInterprete.Text = "Interprete: -"
        lblInterprete.Location = New Point(12, yR)
        lblInterprete.Size = New Size(rw, 18)
        lblInterprete.Font = New Font("Segoe UI", 8)
        pnlRight.Controls.Add(lblInterprete)
        yR += 22

        lblAyantsDroit = New Label()
        lblAyantsDroit.Text = "Ayants droit: 0"
        lblAyantsDroit.Location = New Point(12, yR)
        lblAyantsDroit.Size = New Size(100, 18)
        lblAyantsDroit.Font = New Font("Segoe UI", 8)
        pnlRight.Controls.Add(lblAyantsDroit)

        lblLettrages = New Label()
        lblLettrages.Text = "Lettrages: 0"
        lblLettrages.Location = New Point(120, yR)
        lblLettrages.Size = New Size(100, 18)
        lblLettrages.Font = New Font("Segoe UI", 8)
        pnlRight.Controls.Add(lblLettrages)
        yR += 20

        lblAuteurs = New Label()
        lblAuteurs.Text = "A: 0"
        lblAuteurs.Location = New Point(12, yR)
        lblAuteurs.Size = New Size(60, 18)
        lblAuteurs.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblAuteurs.ForeColor = Color.DarkBlue
        pnlRight.Controls.Add(lblAuteurs)

        lblCompositeurs = New Label()
        lblCompositeurs.Text = "C: 0"
        lblCompositeurs.Location = New Point(80, yR)
        lblCompositeurs.Size = New Size(60, 18)
        lblCompositeurs.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblCompositeurs.ForeColor = Color.DarkBlue
        pnlRight.Controls.Add(lblCompositeurs)

        lblEditeurs = New Label()
        lblEditeurs.Text = "E: 0"
        lblEditeurs.Location = New Point(148, yR)
        lblEditeurs.Size = New Size(60, 18)
        lblEditeurs.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblEditeurs.ForeColor = Color.DarkBlue
        pnlRight.Controls.Add(lblEditeurs)
        yR += 20

        lblNonSACEM = New Label()
        lblNonSACEM.Text = "NON-SACEM: 0"
        lblNonSACEM.Location = New Point(12, yR)
        lblNonSACEM.Size = New Size(150, 18)
        lblNonSACEM.Font = New Font("Segoe UI", 8)
        pnlRight.Controls.Add(lblNonSACEM)

        lblPartsInedites = New Label()
        lblPartsInedites.Text = "Parts inedites: 0"
        lblPartsInedites.Location = New Point(170, yR)
        lblPartsInedites.Size = New Size(150, 18)
        lblPartsInedites.Font = New Font("Segoe UI", 8)
        pnlRight.Controls.Add(lblPartsInedites)
        yR += 28

        ' ── Section : Configuration ───────────────────────────────
        Dim sepCfg As New Panel()
        sepCfg.Location = New Point(12, yR)
        sepCfg.Size = New Size(rw, 1)
        sepCfg.BackColor = Color.FromArgb(200, 215, 235)
        pnlRight.Controls.Add(sepCfg)
        yR += 8

        Dim lblCfgH As New Label()
        lblCfgH.Text = "CONFIGURATION"
        lblCfgH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblCfgH.ForeColor = Color.FromArgb(30, 80, 160)
        lblCfgH.Location = New Point(12, yR)
        lblCfgH.Size = New Size(rw, 18)
        pnlRight.Controls.Add(lblCfgH)
        yR += 22

        pnlRight.Controls.Add(MkCaption("Dossier Templates", 12, yR))
        yR += 14
        txtTemplatesPath = New TextBox()
        txtTemplatesPath.Location = New Point(12, yR)
        txtTemplatesPath.Size = New Size(rw, 20)
        txtTemplatesPath.ReadOnly = True
        txtTemplatesPath.Font = New Font("Segoe UI", 7.5F)
        pnlRight.Controls.Add(txtTemplatesPath)
        yR += 26

        pnlRight.Controls.Add(MkCaption("Dossier Output", 12, yR))
        yR += 14
        txtOutputPath = New TextBox()
        txtOutputPath.Location = New Point(12, yR)
        txtOutputPath.Size = New Size(rw, 20)
        txtOutputPath.ReadOnly = True
        txtOutputPath.Font = New Font("Segoe UI", 7.5F)
        pnlRight.Controls.Add(txtOutputPath)
        yR += 26

        pnlRight.Controls.Add(MkCaption("Template paragraphes", 12, yR))
        yR += 14
        txtParagraphTemplate = New TextBox()
        txtParagraphTemplate.Location = New Point(12, yR)
        txtParagraphTemplate.Size = New Size(rw, 20)
        txtParagraphTemplate.ReadOnly = True
        txtParagraphTemplate.Font = New Font("Segoe UI", 7.5F)
        pnlRight.Controls.Add(txtParagraphTemplate)
        yR += 32

        ' ── Section : Génération ─────────────────────────────────
        Dim sepGen As New Panel()
        sepGen.Location = New Point(12, yR)
        sepGen.Size = New Size(rw, 1)
        sepGen.BackColor = Color.FromArgb(200, 215, 235)
        pnlRight.Controls.Add(sepGen)
        yR += 8

        Dim lblGenH As New Label()
        lblGenH.Text = "GENERATION"
        lblGenH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblGenH.ForeColor = Color.FromArgb(30, 80, 160)
        lblGenH.Location = New Point(12, yR)
        lblGenH.Size = New Size(rw, 18)
        pnlRight.Controls.Add(lblGenH)
        yR += 24

        btnGenerateBDO = MkBtn("Generer BDO", 12, yR, 162, Color.FromArgb(60, 130, 60))
        pnlRight.Controls.Add(btnGenerateBDO)

        btnGenerateContracts = MkBtn("Generer Contrats", 182, yR, 170, Color.FromArgb(60, 130, 60))
        pnlRight.Controls.Add(btnGenerateContracts)
        yR += 36

        btnGenerateAll = MkBtn("TOUT GENERER", 12, yR, rw, Color.FromArgb(20, 100, 20))
        btnGenerateAll.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        pnlRight.Controls.Add(btnGenerateAll)
        yR += 36

        btnOpenOutput = MkBtn("Ouvrir dossier Output", 12, yR, rw, Color.FromArgb(80, 100, 130))
        pnlRight.Controls.Add(btnOpenOutput)
        yR += 36

        ' ── Section : Logs ───────────────────────────────────────
        Dim sepLog As New Panel()
        sepLog.Location = New Point(12, yR)
        sepLog.Size = New Size(rw, 1)
        sepLog.BackColor = Color.FromArgb(200, 215, 235)
        pnlRight.Controls.Add(sepLog)
        yR += 8

        Dim lblLogH As New Label()
        lblLogH.Text = "LOGS"
        lblLogH.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblLogH.ForeColor = Color.FromArgb(30, 80, 160)
        lblLogH.Location = New Point(12, yR)
        lblLogH.Size = New Size(rw, 18)
        pnlRight.Controls.Add(lblLogH)
        yR += 24

        txtLog = New TextBox()
        txtLog.Multiline = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.Location = New Point(12, yR)
        txtLog.Size = New Size(rw, 300)
        txtLog.Font = New Font("Consolas", 8)
        txtLog.ReadOnly = True
        txtLog.BackColor = Color.FromArgb(20, 25, 35)
        txtLog.ForeColor = Color.FromArgb(160, 210, 160)
        pnlRight.Controls.Add(txtLog)
        yR += 308

        btnClearLog = MkBtn("Effacer les logs", 12, yR, rw, Color.FromArgb(100, 80, 80))
        pnlRight.Controls.Add(btnClearLog)
        yR += 40

        pnlRight.Height = yR + 10

        ' ── Événements ───────────────────────────────────────────
        AddHandler Me.Load, AddressOf MainForm_Load
        AddHandler btnSelectJson.Click, AddressOf BtnSelectJson_Click
        AddHandler btnChargerJson.Click, AddressOf BtnChargerJson_Click
        AddHandler btnSauvegarder.Click, AddressOf BtnSauvegarder_Click
        AddHandler btnGenerateBDO.Click, AddressOf BtnGenerateBDO_Click
        AddHandler btnGenerateContracts.Click, AddressOf BtnGenerateContracts_Click
        AddHandler btnGenerateAll.Click, AddressOf BtnGenerateAll_Click
        AddHandler btnOpenOutput.Click, AddressOf BtnOpenOutput_Click
        AddHandler btnClearLog.Click, AddressOf BtnClearLog_Click
        AddHandler btnAjouter.Click, AddressOf BtnAjouter_Click
        AddHandler btnCalculer.Click, AddressOf BtnCalculer_Click
        AddHandler btnGererFiches.Click, AddressOf BtnGererFiches_Click
        AddHandler dgv.MouseUp, AddressOf Dgv_MouseUp
        AddHandler dgv.MouseMove, AddressOf Dgv_MouseMove
        AddHandler dgv.MouseDown, AddressOf Dgv_MouseDown
        AddHandler dgv.DragOver, AddressOf Dgv_DragOver
        AddHandler dgv.DragDrop, AddressOf Dgv_DragDrop
        AddHandler dgv.CellValidating, AddressOf Dgv_CellValidating
        AddHandler dgv.CellBeginEdit, AddressOf Dgv_CellBeginEdit
        AddHandler dgv.DataError, AddressOf Dgv_DataError
        AddHandler dgv.CellDoubleClick, AddressOf Dgv_CellDoubleClick
        AddHandler mnuSupprimer.Click, AddressOf MnuSupprimer_Click
        AddHandler dgv.CellValueChanged, AddressOf Dgv_CellValueChanged
        AddHandler dgv.CurrentCellDirtyStateChanged, AddressOf Dgv_CurrentCellDirtyStateChanged
        AddHandler cbInegalitaire.CheckedChanged, AddressOf CbInegalitaire_CheckedChanged
        AddHandler txtRecherche.TextChanged, AddressOf TxtRecherche_Changed
        AddHandler txtRecherche.GotFocus, AddressOf TxtRecherche_GotFocus
        AddHandler txtRecherche.LostFocus, AddressOf TxtRecherche_LostFocus
        AddHandler lstResultats.DoubleClick, AddressOf LstResultats_DoubleClick
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CHARGEMENT
    ' ─────────────────────────────────────────────────────────────
    Private Sub MainForm_Load(sender As Object, e As EventArgs)
        ' MinSize et SplitterDistance définis après affichage (taille réelle disponible)
        splitMain.Panel1MinSize = 290
        splitMain.Panel2MinSize = 430
        splitMain.SplitterDistance = Math.Max(290, Math.Min(300, splitMain.Width - splitMain.Panel2MinSize - splitMain.SplitterWidth))
        splitRight.Panel1MinSize = 300
        splitRight.Panel2MinSize = 430
        splitRight.SplitterDistance = Math.Max(300, splitRight.Width - 435 - splitRight.SplitterWidth)
        InitDataTable()
        dgv.DataSource = DtDepotCreateur
        ConfigureGridColumns()
        ChargerGoogleSheet()
        BtnClearLog_Click(Nothing, Nothing)
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' HELPERS UI
    ' ─────────────────────────────────────────────────────────────
    Private Function MkCaption(text As String, x As Integer, y As Integer) As Label
        Dim l As New Label()
        l.Text = text
        l.Location = New Point(x, y)
        l.Size = New Size(200, 14)
        l.Font = New Font("Segoe UI", 7.5F)
        l.ForeColor = Color.FromArgb(110, 130, 160)
        Return l
    End Function

    Private Function MkTb(x As Integer, y As Integer, w As Integer) As TextBox
        Dim tb As New TextBox()
        tb.Location = New Point(x, y)
        tb.Size = New Size(w, 23)
        Return tb
    End Function

    Private Function MkBtn(text As String, x As Integer, y As Integer, w As Integer, bg As Color) As Button
        Dim b As New Button()
        b.Text = text
        b.Location = New Point(x, y)
        b.Size = New Size(w, 28)
        b.BackColor = bg
        b.ForeColor = Color.White
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 0
        b.Font = New Font("Segoe UI", 9)
        b.Cursor = Cursors.Hand
        Return b
    End Function

    ' ─────────────────────────────────────────────────────────────
    ' LOGIQUE GÉNÉRATION (ex-MainForm)
    ' ─────────────────────────────────────────────────────────────
    Private Sub InitializePaths()
        If Not Directory.Exists(_templatesDirectory) Then Directory.CreateDirectory(_templatesDirectory)
        If Not Directory.Exists(_outputDirectory) Then Directory.CreateDirectory(_outputDirectory)
        txtTemplatesPath.Text = Path.GetFullPath(_templatesDirectory)
        txtOutputPath.Text = Path.GetFullPath(_outputDirectory)
        txtParagraphTemplate.Text = Path.GetFullPath(_paragraphTemplatePath)
        btnGenerateBDO.Enabled = False
        btnGenerateContracts.Enabled = False
        btnGenerateAll.Enabled = False
        ResetIndicators()
    End Sub

    Private Sub ResetIndicators()
        lblTitre.Text = "Titre: -"
        lblInterprete.Text = "Interprete: -"
        lblAyantsDroit.Text = "Ayants droit: 0"
        lblLettrages.Text = "Lettrages: 0"
        lblAuteurs.Text = "A: 0"
        lblCompositeurs.Text = "C: 0"
        lblEditeurs.Text = "E: 0"
        lblNonSACEM.Text = "NON-SACEM: 0"
        lblNonSACEM.ForeColor = Color.Black
        lblPartsInedites.Text = "Parts inedites: 0"
        lblPartsInedites.ForeColor = Color.Black
    End Sub

    Private Sub BtnSelectJson_Click(sender As Object, e As EventArgs)
        Using dialog As New OpenFileDialog()
            dialog.Title = "Selectionner le fichier JSON SACEM"
            dialog.Filter = "Fichiers JSON (*.json)|*.json|Tous les fichiers (*.*)|*.*"
            If dialog.ShowDialog() = DialogResult.OK Then
                _currentJsonPath = dialog.FileName
                txtJsonPath.Text = _currentJsonPath
                LoadJsonData(_currentJsonPath)
                ApplyRowColors()
                LoadAndValidateForGeneration()
            End If
        End Using
    End Sub

    ''' <summary>Charge le JSON dans le moteur de génération et met à jour les statistiques.</summary>
    Private Sub LoadAndValidateForGeneration()
        Try
            txtLog.AppendText($"{vbCrLf}=== CHARGEMENT JSON ==={vbCrLf}")
            txtLog.AppendText($"Fichier: {Path.GetFileName(_currentJsonPath)}{vbCrLf}")
            _currentData = SACEMJsonReader.LoadFromFile(_currentJsonPath)
            If _currentData Is Nothing Then Throw New Exception("Echec du chargement du JSON")
            Dim validation = SACEMJsonReader.ValidateStructure(_currentData)
            If Not validation.IsValid Then
                txtLog.AppendText($"Attention: {validation.ErrorMessage}{vbCrLf}")
            Else
                txtLog.AppendText($"Structure JSON valide{vbCrLf}")
            End If
            Dim report As String = SACEMJsonReader.GenerateStructureReport(_currentData)
            txtLog.AppendText($"{report}{vbCrLf}")
            lblTitre.Text = $"Titre: {_currentData.Titre}"
            lblInterprete.Text = $"Interprete: {_currentData.Interprete}"
            lblAyantsDroit.Text = $"Ayants droit: {_currentData.AyantsDroit.Count}"
            UpdateDetailedStatistics()
            LoadParagraphTemplates()
            btnGenerateBDO.Enabled = True
            btnGenerateContracts.Enabled = True
            btnGenerateAll.Enabled = True
            txtLog.AppendText($"JSON charge et valide avec succes{vbCrLf}")
        Catch ex As Exception
            txtLog.AppendText($"ERREUR: {ex.Message}{vbCrLf}")
            btnGenerateBDO.Enabled = False
            btnGenerateContracts.Enabled = False
            btnGenerateAll.Enabled = False
            ResetIndicators()
        End Try
    End Sub

    Private Sub UpdateDetailedStatistics()
        If _currentData Is Nothing OrElse _currentData.AyantsDroit Is Nothing Then
            ResetIndicators() : Return
        End If
        Dim countA, countC, countE, countNonSACEM, countPartsInedites As Integer
        Dim lettrages As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim ayantsDroitParLettrage As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        For Each ayant In _currentData.AyantsDroit
            Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
            Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
            Dim societe As String = If(ayant.Identite.SocieteGestion, "SACEM").Trim().ToUpper()
            Select Case role
                Case "A", "AD" : countA += 1
                Case "C", "AR" : countC += 1
                Case "AC" : countA += 1 : countC += 1
                Case "E" : countE += 1
            End Select
            If Not String.IsNullOrEmpty(lettrage) Then
                lettrages.Add(lettrage)
                If Not ayantsDroitParLettrage.ContainsKey(lettrage) Then
                    ayantsDroitParLettrage(lettrage) = New List(Of String)
                End If
                ayantsDroitParLettrage(lettrage).Add(role)
            End If
            If societe <> "SACEM" AndAlso Not String.IsNullOrEmpty(societe) Then countNonSACEM += 1
        Next
        For Each kvp In ayantsDroitParLettrage
            Dim roles As List(Of String) = kvp.Value
            Dim hasAC As Boolean = roles.Any(Function(r) r = "A" OrElse r = "C" OrElse r = "AC" OrElse r = "AR" OrElse r = "AD")
            Dim hasE As Boolean = roles.Any(Function(r) r = "E")
            If hasAC AndAlso Not hasE Then
                countPartsInedites += roles.Where(Function(r) r = "A" OrElse r = "C" OrElse r = "AC" OrElse r = "AR" OrElse r = "AD").Count()
            End If
        Next
        lblLettrages.Text = $"Lettrages: {lettrages.Count}"
        lblAuteurs.Text = $"A: {countA}"
        lblCompositeurs.Text = $"C: {countC}"
        lblEditeurs.Text = $"E: {countE}"
        lblNonSACEM.Text = $"NON-SACEM: {countNonSACEM}"
        lblNonSACEM.ForeColor = If(countNonSACEM > 0, Color.OrangeRed, Color.Green)
        lblNonSACEM.Font = New Font(lblNonSACEM.Font, If(countNonSACEM > 0, FontStyle.Bold, FontStyle.Regular))
        lblPartsInedites.Text = $"Parts inedites: {countPartsInedites}"
        lblPartsInedites.ForeColor = If(countPartsInedites > 0, Color.OrangeRed, Color.Green)
        lblPartsInedites.Font = New Font(lblPartsInedites.Font, If(countPartsInedites > 0, FontStyle.Bold, FontStyle.Regular))
        txtLog.AppendText($"{vbCrLf}=== STATISTIQUES ==={vbCrLf}")
        txtLog.AppendText($"  Lettrages: {lettrages.Count} ({String.Join(", ", lettrages.OrderBy(Function(l) l))}){vbCrLf}")
        txtLog.AppendText($"  A/AD: {countA}  C/AR: {countC}  E: {countE}  NON-SACEM: {countNonSACEM}  Parts inedites: {countPartsInedites}{vbCrLf}")
        If countNonSACEM > 0 Then ShowNonSACEMAlert()
    End Sub

    Private Sub ShowNonSACEMAlert()
        If _currentData Is Nothing OrElse _currentData.AyantsDroit Is Nothing Then Return
        Dim msg As New System.Text.StringBuilder()
        msg.AppendLine("OEUVRE MIXTE DETECTEE")
        msg.AppendLine()
        msg.AppendLine("Ayants droit non membres SACEM :")
        msg.AppendLine()
        Dim acList As New List(Of String)
        Dim eList As New List(Of String)
        For Each ayant In _currentData.AyantsDroit
            Dim societe As String = If(ayant.Identite.SocieteGestion, "").Trim().ToUpper()
            If String.IsNullOrEmpty(societe) OrElse societe = "SACEM" Then Continue For
            Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
            Dim displayName As String = If(ayant.Identite.Type = "Physique",
                $"{ayant.Identite.Prenom} {ayant.Identite.Nom}".Trim(), ayant.Identite.Designation)
            Dim info As String = $"* {displayName} ({ayant.Identite.SocieteGestion})"
            If role = "A" OrElse role = "C" OrElse role = "AC" OrElse role = "AR" OrElse role = "AD" Then
                acList.Add(info)
            ElseIf role = "E" Then
                eList.Add(info)
            End If
        Next
        If acList.Count > 0 Then
            msg.AppendLine("AUTEURS/COMPOSITEURS :")
            For Each s In acList : msg.AppendLine(s) : Next
            msg.AppendLine("-> Ne signeront PAS : CCEOM, CCDAA")
            msg.AppendLine()
        End If
        If eList.Count > 0 Then
            msg.AppendLine("EDITEURS :")
            For Each s In eList : msg.AppendLine(s) : Next
            msg.AppendLine("-> Ne signeront PAS : COED")
            msg.AppendLine()
        End If
        MessageBox.Show(msg.ToString(), "Alerte : Membres NON-SACEM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub LoadParagraphTemplates()
        Try
            txtLog.AppendText($"{vbCrLf}Chargement templates paragraphes...{vbCrLf}")
            If Not File.Exists(_paragraphTemplatePath) Then
                txtLog.AppendText($"Template paragraphes introuvable : {_paragraphTemplatePath}{vbCrLf}")
                Return
            End If
            _paragraphReader = New ParagraphTemplateReader(_paragraphTemplatePath)
            _paragraphReader.LoadTemplates()
            txtLog.AppendText($"Templates paragraphes charges{vbCrLf}")
        Catch ex As Exception
            txtLog.AppendText($"Erreur chargement templates: {ex.Message}{vbCrLf}")
        End Try
    End Sub

    Private Sub BtnGenerateBDO_Click(sender As Object, e As EventArgs)
        Try
            txtLog.AppendText($"{vbCrLf}=== GENERATION BDO ==={vbCrLf}")
            ' DEBUG temporaire — vérifier identité dans _currentData
            If _currentData IsNot Nothing AndAlso _currentData.AyantsDroit IsNot Nothing Then
                For i As Integer = 0 To Math.Min(2, _currentData.AyantsDroit.Count - 1)
                    Dim ay = _currentData.AyantsDroit(i)
                    Dim nomDebug As String = If(ay.Identite IsNot Nothing, $"Type={ay.Identite.Type} Nom={ay.Identite.Nom} Desig={ay.Identite.Designation}", "Identite=Nothing")
                    txtLog.AppendText($"  DEBUG AD{i+1}: {ay.BDO.Id} {ay.BDO.Role} | {nomDebug}{vbCrLf}")
                Next
            End If
            Dim bdoTemplatePath As String = Path.Combine(_templatesDirectory, "Bdo711.pdf")
            If Not File.Exists(bdoTemplatePath) Then
                MessageBox.Show($"Template PDF BDO introuvable:{vbCrLf}{bdoTemplatePath}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Dim outputFileName As String = CleanFileName($"BDO_{_currentData.Titre}_{_currentData.Interprete}.pdf")
            Dim outputPath As String = Path.Combine(_outputDirectory, outputFileName)
            Dim generator As New BDOPdfGenerator(_currentData)
            Dim success As Boolean = generator.Generate(bdoTemplatePath, outputPath)
            For Each logEntry In generator.GenerationLog : txtLog.AppendText($"{logEntry}{vbCrLf}") : Next
            If success Then
                MessageBox.Show($"BDO genere avec succes:{vbCrLf}{outputPath}", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Echec generation BDO. Consultez les logs.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            txtLog.AppendText($"ERREUR: {ex.Message}{vbCrLf}")
            MessageBox.Show($"Erreur:{vbCrLf}{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnGenerateContracts_Click(sender As Object, e As EventArgs)
        Try
            txtLog.AppendText($"{vbCrLf}=== GENERATION CONTRATS ==={vbCrLf}")
            Dim templates As New Dictionary(Of String, String) From {
                {"CCDAA", Path.Combine(_templatesDirectory, "CCDAA_template.docx")},
                {"CCEOM", Path.Combine(_templatesDirectory, "CCEOM_template_univ.docx")},
                {"COED", Path.Combine(_templatesDirectory, "COED_template_univ.docx")}
            }
            Dim missing As New List(Of String)
            For Each kvp In templates
                If Not File.Exists(kvp.Value) Then missing.Add(kvp.Key)
            Next
            If missing.Count > 0 Then
                MessageBox.Show($"Templates manquants:{vbCrLf}{String.Join(vbCrLf, missing)}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            For Each kvp In templates
                Dim outputPath As String = Path.Combine(_outputDirectory, CleanFileName($"{kvp.Key}_{_currentData.Titre}_{_currentData.Interprete}.docx"))
                txtLog.AppendText($"Generation {kvp.Key}...{vbCrLf}")
                Dim contractGenerator As New ContractGenerator(_currentData, _paragraphReader)
                Dim success As Boolean = contractGenerator.Generate(kvp.Value, outputPath, kvp.Key)
                For Each logEntry In contractGenerator.GenerationLog : txtLog.AppendText($"  {logEntry}{vbCrLf}") : Next
                txtLog.AppendText(If(success, $"{kvp.Key} genere OK{vbCrLf}", $"Echec {kvp.Key}{vbCrLf}"))
            Next
            MessageBox.Show("Generation des contrats terminee.", "Termine", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            txtLog.AppendText($"ERREUR: {ex.Message}{vbCrLf}")
            MessageBox.Show($"Erreur:{vbCrLf}{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnGenerateAll_Click(sender As Object, e As EventArgs)
        txtLog.AppendText($"{vbCrLf}{New String("="c, 40)}{vbCrLf}=== GENERATION COMPLETE ==={vbCrLf}{New String("="c, 40)}{vbCrLf}")
        BtnGenerateBDO_Click(sender, e)
        BtnGenerateContracts_Click(sender, e)
        txtLog.AppendText($"{vbCrLf}=== GENERATION COMPLETE TERMINEE ==={vbCrLf}")
    End Sub

    Private Sub BtnClearLog_Click(sender As Object, e As EventArgs)
        txtLog.Clear()
        txtLog.AppendText("=== SACEM GENERATOR ===" & vbCrLf)
        txtLog.AppendText($"Version 1.1 - {DateTime.Now:yyyy-MM-dd}" & vbCrLf & vbCrLf)
    End Sub

    Private Sub BtnOpenOutput_Click(sender As Object, e As EventArgs)
        Try
            If Directory.Exists(_outputDirectory) Then
                Process.Start("explorer.exe", _outputDirectory)
            Else
                MessageBox.Show("Le dossier Output n'existe pas encore.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Impossible d'ouvrir le dossier:{vbCrLf}{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' LOGIQUE ÉDITEUR (ex-JsonEditorForm)
    ' ─────────────────────────────────────────────────────────────
    Private Sub InitDataTable()
        DtDepotCreateur = New DataTable()
        DtDepotCreateur.Columns.Add("Id", GetType(String))
        DtDepotCreateur.Columns.Add("Type", GetType(String))
        DtDepotCreateur.Columns.Add("Designation", GetType(String))
        DtDepotCreateur.Columns.Add("Pseudonyme", GetType(String))
        DtDepotCreateur.Columns.Add("Nom", GetType(String))
        DtDepotCreateur.Columns.Add("Prenom", GetType(String))
        DtDepotCreateur.Columns.Add("Genre", GetType(String))
        DtDepotCreateur.Columns.Add("Nele", GetType(String))
        DtDepotCreateur.Columns.Add("Nea", GetType(String))
        DtDepotCreateur.Columns.Add("SocieteGestion", GetType(String))
        DtDepotCreateur.Columns.Add("FormeJuridique", GetType(String))
        DtDepotCreateur.Columns.Add("Capital", GetType(String))
        DtDepotCreateur.Columns.Add("RCS", GetType(String))
        DtDepotCreateur.Columns.Add("Siren", GetType(String))
        DtDepotCreateur.Columns.Add("GenreRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("PrenomRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("NomRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("FonctionRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("Role", GetType(String))
        DtDepotCreateur.Columns.Add("Lettrage", GetType(String))
        DtDepotCreateur.Columns.Add("COAD_IPI", GetType(String))
        DtDepotCreateur.Columns.Add("PH", GetType(String))
        DtDepotCreateur.Columns.Add("DE", GetType(String))
        DtDepotCreateur.Columns.Add("DR", GetType(String))
        DtDepotCreateur.Columns.Add("Managelic", GetType(String))
        DtDepotCreateur.Columns.Add("Managesub", GetType(String))
        Dim colSig As New DataColumn("Signataire", GetType(Boolean))
        colSig.DefaultValue = True
        DtDepotCreateur.Columns.Add(colSig)
        DtDepotCreateur.Columns.Add("NumVoie", GetType(String))
        DtDepotCreateur.Columns.Add("TypeVoie", GetType(String))
        DtDepotCreateur.Columns.Add("NomVoie", GetType(String))
        DtDepotCreateur.Columns.Add("CP", GetType(String))
        DtDepotCreateur.Columns.Add("Ville", GetType(String))
        DtDepotCreateur.Columns.Add("Pays", GetType(String))
        DtDepotCreateur.Columns.Add("Mail", GetType(String))
        DtDepotCreateur.Columns.Add("Tel", GetType(String))
        DtDepotCreateur.Columns.Add("_EditeurDefaut", GetType(String))
        DtDepotCreateur.Columns.Add("_Orphelin", GetType(Boolean))
    End Sub

    Private Sub ConfigureGridColumns()
        For Each col As DataGridViewColumn In dgv.Columns
            col.Visible = False
        Next
        Dim visible As New Dictionary(Of String, Integer) From {
            {"Id", 70}, {"Designation", 250}, {"Role", 45}, {"Lettrage", 60},
            {"SocieteGestion", 90}, {"Signataire", 70}, {"COAD_IPI", 130},
            {"PH", 60}, {"DE", 70}, {"DR", 70}
        }
        For Each kvp In visible
            If dgv.Columns.Contains(kvp.Key) Then
                dgv.Columns(kvp.Key).Visible = True
                dgv.Columns(kvp.Key).Width = kvp.Value
            End If
        Next
        dgv.Columns("SocieteGestion").HeaderText = "Societe"
        dgv.Columns("COAD_IPI").HeaderText = "COAD/IPI"
        dgv.Columns("Signataire").HeaderText = "Signataire"
        dgv.Columns("PH").HeaderText = "PH %"
        If dgv.Columns.Contains("DE") Then dgv.Columns("DE").HeaderText = "DEP %" : dgv.Columns("DE").ReadOnly = True
        If dgv.Columns.Contains("DR") Then dgv.Columns("DR").HeaderText = "DR %" : dgv.Columns("DR").ReadOnly = True
        SetupMoraleComboColumn("Managelic", 150)
        SetupMoraleComboColumn("Managesub", 150)
    End Sub

    Private Sub SetupMoraleComboColumn(colName As String, width As Integer)
        If dgv.Columns.Contains(colName) Then dgv.Columns.Remove(colName)
        Dim cbCol As New DataGridViewComboBoxColumn()
        cbCol.Name = colName
        cbCol.HeaderText = colName
        cbCol.DataPropertyName = colName
        cbCol.Width = width
        cbCol.Visible = True
        cbCol.FlatStyle = FlatStyle.Flat
        cbCol.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
        cbCol.Items.Add("")
        If DtPersonMor IsNot Nothing Then
            For Each row As DataRow In DtPersonMor.Rows
                Dim desig As String = row(0).ToString().Trim()
                If Not String.IsNullOrEmpty(desig) AndAlso Not cbCol.Items.Contains(desig) Then
                    cbCol.Items.Add(desig)
                End If
            Next
        End If
        dgv.Columns.Add(cbCol)
    End Sub

    Private Sub RefreshMoraleColumns()
        SetupMoraleComboColumn("Managelic", 150)
        SetupMoraleComboColumn("Managesub", 150)
    End Sub

    Public Sub RefreshDeclarationFormat()
        Dim currentDecl As String = cbDeclaration.Text
        Dim currentFmt As String = cbFormat.Text
        cbDeclaration.Items.Clear()
        cbFormat.Items.Clear()
        cbDeclaration.Items.Add("")
        cbFormat.Items.Add("")
        For Each row As DataRow In DtDepotCreateur.Rows
            Dim role As String = row("Role").ToString().Trim().ToUpper()
            If MoteurRepartition.IsEditeur(role) Then
                Dim desig As String = row("Designation").ToString().Trim()
                If Not String.IsNullOrEmpty(desig) Then
                    If Not cbDeclaration.Items.Contains(desig) Then cbDeclaration.Items.Add(desig)
                    If Not cbFormat.Items.Contains(desig) Then cbFormat.Items.Add(desig)
                End If
            End If
        Next
        cbDeclaration.Text = currentDecl
        cbFormat.Text = currentFmt
    End Sub

    Private Sub ChargerGoogleSheet()
        Dim localPath As String = PersonnesForm.DefaultXlsxPath
        Try
            If File.Exists(localPath) Then
                lblStatut.Text = "Chargement de la base de donnees..."
                Application.DoEvents()
                DtPersonPhy = LoadSheetXlsxLocal(localPath, "PERSONNEPHYSIQUE")
                DtPersonMor = LoadSheetXlsxLocal(localPath, "PERSONNEMORALE")
                PopulateComboBox()
                RefreshMoraleColumns()
                lblStatut.Text = DtPersonPhy.Rows.Count & " physiques, " & DtPersonMor.Rows.Count & " morales charges."
            Else
                lblStatut.Text = "Fichier BDD introuvable : " & localPath
                DtPersonPhy = New DataTable()
                DtPersonMor = New DataTable()
            End If
        Catch ex As Exception
            lblStatut.Text = "Erreur chargement BDD : " & ex.Message
            DtPersonPhy = New DataTable()
            DtPersonMor = New DataTable()
        End Try
    End Sub

    Private Function LoadSheetXlsxLocal(path As String, sheetName As String) As DataTable
        Dim dt As New DataTable()
        Using pkg As New ExcelPackage(New FileInfo(path))
            Dim ws = pkg.Workbook.Worksheets(sheetName)
            If ws Is Nothing OrElse ws.Dimension Is Nothing Then Return dt
            For col = 1 To ws.Dimension.Columns
                dt.Columns.Add(ws.Cells(1, col).Text)
            Next
            For row = 2 To ws.Dimension.Rows
                Dim nr = dt.NewRow()
                For col = 1 To ws.Dimension.Columns
                    nr(col - 1) = ws.Cells(row, col).Text
                Next
                Dim isEmpty As Boolean = True
                For col = 0 To dt.Columns.Count - 1
                    If Not String.IsNullOrEmpty(nr(col).ToString()) Then isEmpty = False : Exit For
                Next
                If Not isEmpty Then dt.Rows.Add(nr)
            Next
        End Using
        Return dt
    End Function

    Private Sub BtnGererFiches_Click(sender As Object, e As EventArgs)
        Dim xlsxPath As String = PersonnesForm.DefaultXlsxPath
        Using f As New PersonnesForm(xlsxPath)
            f.ShowDialog()
        End Using
        ChargerGoogleSheet()
        SyncGridAvecSheet()
    End Sub

    Private Sub SyncGridAvecSheet()
        Dim mapPhyXls As String() = {"Pseudonyme", "Nom", "Prenom", "Genre", "SocieteGestion", "Num de voie", "Type de voie", "Nom de voie", "CP", "Ville", "Mail", "Tel", "Date de naissance", "Lieu de naissance", "N Secu"}
        Dim mapPhyGrd As String() = {"Pseudonyme", "Nom", "Prenom", "Genre", "SocieteGestion", "NumVoie", "TypeVoie", "NomVoie", "CP", "Ville", "Mail", "Tel", "Nele", "Nea", "NSecu"}
        Dim mapMorXls As String() = {"Designation", "SocieteGestion", "Forme Juridique", "Capital", "RCS", "Siren", "Num de voie", "Type de voie", "Nom de voie", "CP", "Ville", "Mail", "Tel", "Prenom representant", "Nom representant", "Fonction representant"}
        Dim mapMorGrd As String() = {"Designation", "SocieteGestion", "FormeJuridique", "Capital", "RCS", "Siren", "NumVoie", "TypeVoie", "NomVoie", "CP", "Ville", "Mail", "Tel", "PrenomRepresentant", "NomRepresentant", "FonctionRepresentant"}
        For Each gridRow As DataRow In DtDepotCreateur.Rows
            Dim tp As String = gridRow("Type").ToString()
            Dim sheetRow As DataRow = Nothing
            If tp = "Physique" AndAlso DtPersonPhy IsNot Nothing Then
                Dim nom As String = gridRow("Nom").ToString().Trim().ToUpper()
                Dim prenom As String = gridRow("Prenom").ToString().Trim().ToUpper()
                Dim pseudo As String = gridRow("Pseudonyme").ToString().Trim().ToUpper()
                For Each r As DataRow In DtPersonPhy.Rows
                    If (Not String.IsNullOrEmpty(nom) AndAlso SafeStr(r, "Nom").Trim().ToUpper() = nom AndAlso SafeStr(r, "Prenom").Trim().ToUpper() = prenom) OrElse
                       (Not String.IsNullOrEmpty(pseudo) AndAlso SafeStr(r, "Pseudonyme").Trim().ToUpper() = pseudo) Then
                        sheetRow = r : Exit For
                    End If
                Next
                If sheetRow IsNot Nothing Then
                    For k As Integer = 0 To mapPhyXls.Length - 1
                        If DtDepotCreateur.Columns.Contains(mapPhyGrd(k)) Then
                            gridRow(mapPhyGrd(k)) = SafeStr(sheetRow, mapPhyXls(k))
                        End If
                    Next
                    gridRow("COAD_IPI") = GetCOADIPI(sheetRow)
                    Dim nomXls As String = SafeStr(sheetRow, "Nom").Trim()
                    Dim prenomXls As String = SafeStr(sheetRow, "Prenom").Trim()
                    Dim pseudoXls As String = SafeStr(sheetRow, "Pseudonyme").Trim()
                    gridRow("Designation") = If(Not String.IsNullOrEmpty(pseudoXls),
                        nomXls & " " & prenomXls & " / " & pseudoXls,
                        (nomXls & " " & prenomXls).Trim())
                End If
            ElseIf tp = "Moral" AndAlso DtPersonMor IsNot Nothing Then
                Dim desigUp As String = gridRow("Designation").ToString().Trim().ToUpper()
                For Each r As DataRow In DtPersonMor.Rows
                    If SafeStr(r, "Designation").Trim().ToUpper() = desigUp Then sheetRow = r : Exit For
                Next
                If sheetRow IsNot Nothing Then
                    For k As Integer = 0 To mapMorXls.Length - 1
                        If DtDepotCreateur.Columns.Contains(mapMorGrd(k)) Then
                            gridRow(mapMorGrd(k)) = SafeStr(sheetRow, mapMorXls(k))
                        End If
                    Next
                    gridRow("COAD_IPI") = GetCOADIPI(sheetRow)
                End If
            End If
        Next
        dgv.Refresh()
        lblStatut.Text = "Grille synchronisee avec les fiches."
    End Sub

    Private Sub PopulateComboBox()
        cbPersonnes.Items.Clear()
        Dim items As New List(Of String)()
        If DtPersonPhy IsNot Nothing Then
            For Each row As DataRow In DtPersonPhy.Rows
                Dim pseudo As String = SafeStr(row, "Pseudonyme")
                Dim nom As String = SafeStr(row, "Nom")
                Dim prenom As String = SafeStr(row, "Prenom")
                If Not String.IsNullOrEmpty(pseudo) OrElse Not String.IsNullOrEmpty(nom & prenom) Then
                    items.Add($"{pseudo} / {nom} / {prenom}")
                End If
            Next
        End If
        If DtPersonMor IsNot Nothing Then
            For Each row As DataRow In DtPersonMor.Rows
                Dim desig As String = SafeStr(row, "Designation")
                If Not String.IsNullOrEmpty(desig) Then items.Add($"(Morale) {desig}")
            Next
        End If
        items.Sort()
        cbPersonnes.Items.AddRange(items.ToArray())
    End Sub

    Private Sub BtnAjouter_Click(sender As Object, e As EventArgs)
        Dim sel As String = If(cbPersonnes.SelectedItem IsNot Nothing, cbPersonnes.SelectedItem.ToString(), "")
        If String.IsNullOrEmpty(sel) Then
            MessageBox.Show("Selectionnez une personne.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        RafraichirBDD()
        If sel.StartsWith("(Morale) ") Then
            AjouterPersonneMorale(sel.Substring(9).Trim())
        Else
            AjouterPersonnePhysique(sel)
        End If
        UpdateLettrages()
        ApplyRowColors()
        RefreshDeclarationFormat()
        dgv.Refresh()
    End Sub

    Private Sub RafraichirBDD()
        Dim localPath As String = PersonnesForm.DefaultXlsxPath
        If Not File.Exists(localPath) Then Return
        Try
            DtPersonPhy = LoadSheetXlsxLocal(localPath, "PERSONNEPHYSIQUE")
            DtPersonMor = LoadSheetXlsxLocal(localPath, "PERSONNEMORALE")
        Catch
        End Try
    End Sub

    Private Sub AjouterPersonneMorale(designation As String)
        If DtPersonMor Is Nothing Then Return
        Dim foundRow As DataRow = DtPersonMor.AsEnumerable().FirstOrDefault(Function(r) SafeStr(r, "Designation").Trim().ToUpper() = designation.Trim().ToUpper())
        If foundRow Is Nothing Then Return
        If DtDepotCreateur.Rows.Count = 0 Then
            MessageBox.Show("Ajoutez d'abord un createur (A ou C) avant un editeur.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim createurs As New List(Of String)(
            DtDepotCreateur.AsEnumerable().Where(Function(r) r("Role").ToString() = "A" OrElse r("Role").ToString() = "C").Select(Function(r) r("Designation").ToString()))
        Using f As New FormCreateurs(createurs)
            If f.ShowDialog() <> DialogResult.OK Then Return
            For Each crea As String In f.SelectedCreateurs
                Dim creaRow As DataRow = DtDepotCreateur.Select($"Designation = '{crea.Replace("'", "''")}'").FirstOrDefault()
                If creaRow Is Nothing Then Continue For
                If String.IsNullOrEmpty(creaRow("Lettrage").ToString()) Then creaRow("Lettrage") = GetNextLetter()
                Dim nr As DataRow = DtDepotCreateur.NewRow()
                nr("Id") = SafeStr(foundRow, "Id")
                nr("Type") = "Moral"
                nr("Designation") = designation
                nr("Role") = "E"
                nr("Lettrage") = creaRow("Lettrage")
                nr("SocieteGestion") = SafeStr(foundRow, "SocieteGestion", "SACEM")
                nr("Signataire") = True
                CopierAdresseContact(nr, foundRow)
                CopierInfosMorale(nr, foundRow)
                DtDepotCreateur.Rows.Add(nr)
                MajEditeurDansXlsx(crea, designation)
            Next
        End Using
    End Sub

    Private Sub MajEditeurDansXlsx(creaDesignation As String, editeurDesignation As String)
        If DtPersonPhy Is Nothing Then Return
        Dim phyRow As DataRow = DtPersonPhy.AsEnumerable().FirstOrDefault(
            Function(r)
                Dim pseudo As String = SafeStr(r, "Pseudonyme").Trim()
                Dim nom As String = SafeStr(r, "Nom").Trim()
                Dim prenom As String = SafeStr(r, "Prenom").Trim()
                Dim desig As String = If(Not String.IsNullOrEmpty(pseudo), pseudo, (nom & " " & prenom).Trim())
                Return desig.ToUpper() = creaDesignation.Trim().ToUpper()
            End Function)
        If phyRow Is Nothing Then Return
        Dim current As String = If(DtPersonPhy.Columns.Contains("Editeur"), SafeStr(phyRow, "Editeur"), "")
        Dim editeurs As New List(Of String)(
            current.Split(";"c).Select(Function(s) s.Trim()).Where(Function(s) Not String.IsNullOrEmpty(s)))
        If Not editeurs.Any(Function(e) e.ToUpper() = editeurDesignation.Trim().ToUpper()) Then
            editeurs.Add(editeurDesignation.Trim())
            phyRow("Editeur") = String.Join(";", editeurs)
            SauvegarderXlsxSilencieux()
        End If
    End Sub

    Private Sub SauvegarderXlsxSilencieux()
        Try
            Dim localPath As String = PersonnesForm.DefaultXlsxPath
            If Not File.Exists(localPath) Then Return
            Using pkg As New ExcelPackage(New FileInfo(localPath))
                Dim existing = pkg.Workbook.Worksheets("PERSONNEPHYSIQUE")
                If existing IsNot Nothing Then pkg.Workbook.Worksheets.Delete(existing)
                Dim ws = pkg.Workbook.Worksheets.Add("PERSONNEPHYSIQUE")
                Dim cols As String() = PersonnesForm.ColsPhy
                For c = 0 To cols.Length - 1
                    ws.Cells(1, c + 1).Value = cols(c)
                    ws.Cells(1, c + 1).Style.Font.Bold = True
                Next
                For r = 0 To DtPersonPhy.Rows.Count - 1
                    For c = 0 To cols.Length - 1
                        If DtPersonPhy.Columns.Contains(cols(c)) Then
                            ws.Cells(r + 2, c + 1).Value = DtPersonPhy.Rows(r)(cols(c)).ToString()
                        End If
                    Next
                Next
                pkg.Save()
            End Using
        Catch
        End Try
    End Sub

    Private Sub AjouterPersonnePhysique(selectedValue As String)
        If DtPersonPhy Is Nothing Then Return
        Dim parts() As String = selectedValue.Split("/"c)
        If parts.Length < 3 Then Return
        Dim pseudo As String = parts(0).Trim()
        Dim nom As String = parts(1).Trim()
        Dim prenom As String = parts(2).Trim()
        Dim foundRow As DataRow = Nothing
        For Each r As DataRow In DtPersonPhy.Rows
            If (r("Nom").ToString().Trim() = nom AndAlso SafeStr(r, "Prenom").Trim() = prenom) OrElse
               (Not String.IsNullOrEmpty(pseudo) AndAlso r("Pseudonyme").ToString().Trim() = pseudo) Then
                foundRow = r : Exit For
            End If
        Next
        If foundRow Is Nothing Then Return
        Dim genre As String = If(cbGenre.SelectedItem IsNot Nothing, cbGenre.SelectedItem.ToString(), cbGenre.Text)
        Dim role As String = SafeStr(foundRow, "Role")
        role = AjusterRole(role, genre)
        If String.IsNullOrEmpty(role) Then
            Dim roles As New List(Of String)(RolesPossibles(genre))
            Using f As New FormRoles(roles)
                If f.ShowDialog() <> DialogResult.OK Then Return
                role = f.SelectedRole
            End Using
        End If
        Dim nr As DataRow = DtDepotCreateur.NewRow()
        nr("Id") = SafeStr(foundRow, "Id")
        nr("Type") = "Physique"
        nr("Designation") = BuildDesignation(foundRow)
        nr("Pseudonyme") = SafeStr(foundRow, "Pseudonyme")
        nr("Nom") = SafeStr(foundRow, "Nom")
        nr("Prenom") = SafeStr(foundRow, "Prenom")
        nr("Genre") = SafeStr(foundRow, "Genre")
        nr("SocieteGestion") = SafeStr(foundRow, "SocieteGestion", "SACEM")
        nr("Role") = role
        nr("COAD_IPI") = GetCOADIPI(foundRow)
        nr("Signataire") = True
        CopierAdresseContact(nr, foundRow)
        nr("_EditeurDefaut") = If(foundRow.Table.Columns.Contains("Editeur"), SafeStr(foundRow, "Editeur"), "")
        Dim editeurDefaut As String = If(foundRow.Table.Columns.Contains("Editeur"), SafeStr(foundRow, "Editeur"), "")
        DtDepotCreateur.Rows.Add(nr)
        nr("Lettrage") = GetNextLetter()
        If Not String.IsNullOrEmpty(editeurDefaut) Then
            Dim entrees() As String = editeurDefaut.Split(";"c)
            Dim editeurIds As New List(Of String)()
            Dim partsExplicites As New Dictionary(Of String, Double)()
            For Each entree As String In entrees
                Dim e As String = entree.Trim()
                If String.IsNullOrEmpty(e) Then Continue For
                Dim tokens() As String = e.Split(":"c)
                Dim eid As String = tokens(0).Trim()
                editeurIds.Add(eid)
                If tokens.Length > 1 Then
                    Dim p As Double
                    If Double.TryParse(tokens(1).Trim(), p) Then partsExplicites(eid) = p
                End If
            Next
            Dim nbEds As Integer = editeurIds.Count
            For Each eid As String In editeurIds
                If Not partsExplicites.ContainsKey(eid) Then partsExplicites(eid) = Math.Round(100.0 / nbEds, 2)
            Next
            For Each eid As String In editeurIds
                AjouterEditeurParDefaut(eid, nr, partsExplicites(eid) / 100.0)
            Next
        Else
            If DtDepotCreateur.AsEnumerable().Any(Function(r) MoteurRepartition.IsEditeur(r("Role").ToString())) Then
                If MessageBox.Show("Voulez-vous etre Editeur A Compte d'Auteur (EAC) ?",
                                   "EAC ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Dim nrEAC As DataRow = DtDepotCreateur.NewRow()
                    nrEAC("Id") = nr("Id")
                    nrEAC("Type") = "Physique"
                    nrEAC("Designation") = nr("Designation").ToString() & " (EAC)"
                    nrEAC("Nom") = nr("Nom")
                    nrEAC("Prenom") = nr("Prenom")
                    nrEAC("Role") = "AEC"
                    nrEAC("Lettrage") = nr("Lettrage")
                    nrEAC("SocieteGestion") = nr("SocieteGestion")
                    nrEAC("Signataire") = True
                    CopierAdresseContact(nrEAC, foundRow)
                    DtDepotCreateur.Rows.Add(nrEAC)
                End If
            End If
        End If
    End Sub

    Private Sub AjouterEditeurParDefaut(editeurId As String, creaRow As DataRow,
                                        Optional quotePartCoed As Double = 1.0)
        Dim phProv As Double = Math.Round(quotePartCoed * 100.0, 4)
        Dim phStr As String = phProv.ToString(Globalization.CultureInfo.InvariantCulture)
        If editeurId = "EAC" Then
            Dim nrEAC2 As DataRow = DtDepotCreateur.NewRow()
            nrEAC2("Id") = creaRow("Id")
            nrEAC2("Type") = "Physique"
            nrEAC2("Designation") = creaRow("Designation").ToString() & " (EAC)"
            nrEAC2("Nom") = creaRow("Nom")
            nrEAC2("Prenom") = creaRow("Prenom")
            nrEAC2("Role") = "AEC"
            nrEAC2("Lettrage") = creaRow("Lettrage")
            nrEAC2("PH") = phStr
            nrEAC2("SocieteGestion") = creaRow("SocieteGestion")
            nrEAC2("Signataire") = True
            DtDepotCreateur.Rows.Add(nrEAC2)
            Return
        End If
        Dim morRow As DataRow = Nothing
        If DtPersonMor IsNot Nothing Then
            If editeurId.StartsWith("M") AndAlso DtPersonMor.Columns.Contains("Id") Then
                morRow = DtPersonMor.AsEnumerable().FirstOrDefault(
                    Function(r) SafeStr(r, "Id").Trim().ToUpper() = editeurId.Trim().ToUpper())
            End If
            If morRow Is Nothing Then
                morRow = DtPersonMor.AsEnumerable().FirstOrDefault(
                    Function(r) r("Designation").ToString().Trim().ToUpper() = editeurId.Trim().ToUpper())
            End If
        End If
        If morRow Is Nothing Then Return
        Dim nrEMor As DataRow = DtDepotCreateur.NewRow()
        nrEMor("Type") = "Moral"
        nrEMor("Designation") = SafeStr(morRow, "Designation")
        nrEMor("Id") = SafeStr(morRow, "Id")
        nrEMor("Role") = "E"
        nrEMor("Lettrage") = creaRow("Lettrage")
        nrEMor("PH") = phStr
        nrEMor("SocieteGestion") = SafeStr(morRow, "SocieteGestion", "SACEM")
        nrEMor("Signataire") = True
        CopierInfosMorale(nrEMor, morRow)
        CopierAdresseContact(nrEMor, morRow)
        DtDepotCreateur.Rows.Add(nrEMor)
    End Sub

    Private Function GetParamsOeuvre() As MoteurRepartition.ParamsOeuvre
        Dim p As New MoteurRepartition.ParamsOeuvre()
        Dim genre As String = If(cbGenre.SelectedItem IsNot Nothing,
                                  cbGenre.SelectedItem.ToString(), cbGenre.Text).ToLower()
        If genre.Contains("texte") OrElse genre.Contains("chronique") OrElse
           genre.Contains("poeme") OrElse genre.Contains("sketch") OrElse
           genre.Contains("billet") Then
            p.TypeOeuvre = MoteurRepartition.TypeOeuvre.LitteraireSeule
        ElseIf genre.Contains("instrumental") Then
            p.TypeOeuvre = MoteurRepartition.TypeOeuvre.MusiqueSeule
        Else
            p.TypeOeuvre = MoteurRepartition.TypeOeuvre.ParolesEtMusique
        End If
        p.EstEditee = DtDepotCreateur.AsEnumerable().Any(Function(r) MoteurRepartition.IsEditeur(r("Role").ToString()))
        p.EstDomainePublic = False
        p.EstFilmOuSymphonique = genre.Contains("film") OrElse genre.Contains("symphonique")
        p.Inegalitaire = cbInegalitaire.Checked
        Return p
    End Function

    Private Sub AppelerMoteur()
        Try
            Dim params As MoteurRepartition.ParamsOeuvre = GetParamsOeuvre()
            MoteurRepartition.RecalculerPHApresAjout(DtDepotCreateur, params)
            MoteurRepartition.Calculer(DtDepotCreateur, params)
        Catch ex As Exception
            lblStatut.Text = "Erreur moteur : " & ex.Message
        End Try
    End Sub

    Private Sub BtnCalculer_Click(sender As Object, e As EventArgs)
        Try
            Dim rows = DtDepotCreateur.AsEnumerable().ToList()
            Dim countA As Integer = rows.Where(Function(r) r("Role").ToString() = "A").Count()
            Dim countC As Integer = rows.Where(Function(r) r("Role").ToString() = "C").Count()
            Dim countE As Integer = rows.Where(Function(r) MoteurRepartition.IsEditeur(r("Role").ToString())).Count()
            Dim inegal As Boolean = cbInegalitaire.Checked
            For Each row As DataRow In rows
                Dim role As String = row("Role").ToString()
                Dim de As Decimal = 0
                Dim dr As Decimal = 0
                If Not inegal Then
                    If countE > 0 Then
                        If countA > 0 AndAlso countC > 0 Then
                            If role = "A" Then de = Math.Round(100D / 3 / countA, 4) : dr = de
                            If role = "C" Then de = Math.Round(100D / 3 / countC, 4) : dr = de
                            If MoteurRepartition.IsEditeur(role) Then de = Math.Round(100D / 3 / countE, 4) : dr = de
                        ElseIf countA > 0 Then
                            If role = "A" Then de = Math.Round(200D / 3 / countA, 4) : dr = Math.Round(50D / countA, 4)
                            If MoteurRepartition.IsEditeur(role) Then de = Math.Round(100D / 3 / countE, 4) : dr = Math.Round(50D / countE, 4)
                        ElseIf countC > 0 Then
                            If role = "C" Then de = Math.Round(200D / 3 / countC, 4) : dr = Math.Round(50D / countC, 4)
                            If MoteurRepartition.IsEditeur(role) Then de = Math.Round(100D / 3 / countE, 4) : dr = Math.Round(50D / countE, 4)
                        End If
                    Else
                        If role = "A" Then de = Math.Round(100D / countA, 4) : dr = de
                        If role = "C" Then de = Math.Round(100D / countC, 4) : dr = de
                    End If
                End If
                row("DE") = de.ToString(Globalization.CultureInfo.InvariantCulture)
                row("DR") = dr.ToString(Globalization.CultureInfo.InvariantCulture)
            Next
            UpdateLettrages()
            dgv.Refresh()
            lblStatut.Text = "Pourcentages calcules."
        Catch ex As Exception
            MessageBox.Show($"Erreur calcul : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnSauvegarder_Click(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(txtTitre.Text.Trim()) Then
            MessageBox.Show("Le titre est obligatoire.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Using sfd As New SaveFileDialog()
            sfd.Title = "Enregistrer le fichier JSON SACEM"
            sfd.Filter = "Fichiers JSON (*.json)|*.json"
            sfd.FileName = CleanFileName(txtTitre.Text.Trim()) & ".json"
            If sfd.ShowDialog() <> DialogResult.OK Then Return
            Dim obj As New JObject()
            obj("Titre") = txtTitre.Text.Trim()
            obj("SousTitre") = txtSousTitre.Text.Trim()
            obj("Interprete") = txtInterprete.Text.Trim()
            obj("Duree") = txtDuree.Text.Trim()
            obj("Genre") = If(cbGenre.SelectedItem IsNot Nothing, cbGenre.SelectedItem.ToString(), cbGenre.Text)
            obj("Date") = dtDate.Value.ToString("dd/MM/yyyy")
            obj("ISWC") = txtISWC.Text.Trim()
            obj("Lieu") = If(cbLieu.SelectedItem IsNot Nothing, cbLieu.SelectedItem.ToString(), cbLieu.Text)
            obj("Territoire") = If(cbTerritoire.SelectedItem IsNot Nothing, cbTerritoire.SelectedItem.ToString(), cbTerritoire.Text)
            obj("Arrangement") = If(cbArrangement.SelectedItem IsNot Nothing, cbArrangement.SelectedItem.ToString(), cbArrangement.Text)
            obj("Inegalitaire") = If(cbInegalitaire.Checked, "TRUE", "FALSE")
            obj("Declaration") = cbDeclaration.Text.Trim()
            obj("Format") = cbFormat.Text.Trim()
            obj("Faita") = txtFaita.Text.Trim()
            obj("Faitle") = dtFaitle.Value.ToString("dd/MM/yyyy")
            obj("Commentaire") = ""
            Dim arr As New JArray()
            For Each row As DataRow In DtDepotCreateur.Rows
                Dim ad As New JObject()

                ' ── BDO (clés de répartition) ──
                ad("Id") = row("Id").ToString().Trim()
                ad("Role") = row("Role").ToString()
                ad("Lettrage") = row("Lettrage").ToString()
                ad("PH") = row("PH").ToString()
                ad("DE") = row("DE").ToString()
                ad("DR") = row("DR").ToString()
                ad("Managelic") = row("Managelic").ToString()
                ad("Managesub") = row("Managesub").ToString()
                Dim sig As Boolean = True
                Boolean.TryParse(row("Signataire").ToString(), sig)
                ad("Signataire") = sig

                ' ── Identité ──
                Dim identite As New JObject()
                identite("Type") = row("Type").ToString()
                identite("Designation") = row("Designation").ToString()
                identite("Pseudonyme") = row("Pseudonyme").ToString()
                identite("Nom") = row("Nom").ToString()
                identite("Prenom") = row("Prenom").ToString()
                identite("Genre") = row("Genre").ToString()
                identite("Nele") = row("Nele").ToString()
                identite("Nea") = row("Nea").ToString()
                identite("SocieteGestion") = row("SocieteGestion").ToString()
                ' Morale
                identite("FormeJuridique") = row("FormeJuridique").ToString()
                identite("Capital") = row("Capital").ToString()
                identite("RCS") = row("RCS").ToString()
                identite("Siren") = row("Siren").ToString()
                identite("GenreRepresentant") = row("GenreRepresentant").ToString()
                identite("PrenomRepresentant") = row("PrenomRepresentant").ToString()
                identite("NomRepresentant") = row("NomRepresentant").ToString()
                identite("FonctionRepresentant") = row("FonctionRepresentant").ToString()
                ad("Identite") = identite

                ' ── Adresse ──
                Dim adresse As New JObject()
                adresse("NumVoie") = row("NumVoie").ToString()
                adresse("TypeVoie") = row("TypeVoie").ToString()
                adresse("NomVoie") = row("NomVoie").ToString()
                adresse("CP") = row("CP").ToString()
                adresse("Ville") = row("Ville").ToString()
                adresse("Pays") = row("Pays").ToString()
                ad("Adresse") = adresse

                ' ── Contact ──
                Dim contact As New JObject()
                contact("Mail") = row("Mail").ToString()
                contact("Tel") = row("Tel").ToString()
                ad("Contact") = contact

                ' ── COAD/IPI ──
                ad("COAD_IPI") = row("COAD_IPI").ToString()

                arr.Add(ad)
            Next
            obj("AyantsDroit") = arr
            File.WriteAllText(sfd.FileName, obj.ToString(Newtonsoft.Json.Formatting.Indented), System.Text.Encoding.UTF8)
            _currentJsonPath = sfd.FileName
            txtJsonPath.Text = sfd.FileName
            lblStatut.Text = "JSON sauvegarde : " & sfd.FileName
            ' Recharger via SACEMJsonReader pour que la generation soit immediatement a jour
            LoadAndValidateForGeneration()
            MessageBox.Show("JSON sauvegarde avec succes !" & vbCrLf & sfd.FileName, "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub

    Private Sub BtnChargerJson_Click(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Fichiers JSON (*.json)|*.json"
            ofd.Title = "Charger un fichier JSON SACEM"
            If ofd.ShowDialog() = DialogResult.OK Then
                _currentJsonPath = ofd.FileName
                txtJsonPath.Text = ofd.FileName
                LoadJsonData(ofd.FileName)
                ApplyRowColors()
                LoadAndValidateForGeneration()
            End If
        End Using
    End Sub

    Private Sub LoadJsonData(filePath As String)
        Try
            Dim sacemData As SACEMData = SACEMJsonReader.LoadFromFile(filePath)
            Dim obj As JObject = sacemData.RawData
            txtTitre.Text = If(obj("Titre") IsNot Nothing, obj("Titre").ToString(), "")
            txtSousTitre.Text = If(obj("SousTitre") IsNot Nothing, obj("SousTitre").ToString(), "")
            txtInterprete.Text = If(obj("Interprete") IsNot Nothing, obj("Interprete").ToString(), "")
            txtDuree.Text = If(obj("Duree") IsNot Nothing, obj("Duree").ToString(), "")
            Dim genre As String = If(obj("Genre") IsNot Nothing, obj("Genre").ToString(), "")
            If cbGenre.Items.Contains(genre) Then cbGenre.SelectedItem = genre Else cbGenre.Text = genre
            Dim d As DateTime
            If DateTime.TryParseExact(If(obj("Date") IsNot Nothing, obj("Date").ToString(), ""), "dd/MM/yyyy",
                                      Globalization.CultureInfo.InvariantCulture,
                                      Globalization.DateTimeStyles.None, d) Then dtDate.Value = d
            cbLieu.Text = If(obj("Lieu") IsNot Nothing, obj("Lieu").ToString(), "")
            cbTerritoire.Text = If(obj("Territoire") IsNot Nothing, obj("Territoire").ToString(), "")
            cbArrangement.Text = If(obj("Arrangement") IsNot Nothing, obj("Arrangement").ToString(), "")
            txtISWC.Text = If(obj("ISWC") IsNot Nothing, obj("ISWC").ToString(), "")
            cbDeclaration.Text = If(obj("Declaration") IsNot Nothing, obj("Declaration").ToString(), "")
            cbFormat.Text = If(obj("Format") IsNot Nothing, obj("Format").ToString(), "")
            txtFaita.Text = If(obj("Faita") IsNot Nothing, obj("Faita").ToString(), "")
            If DateTime.TryParseExact(If(obj("Faitle") IsNot Nothing, obj("Faitle").ToString(), ""), "dd/MM/yyyy",
                                      Globalization.CultureInfo.InvariantCulture,
                                      Globalization.DateTimeStyles.None, d) Then dtFaitle.Value = d
            cbInegalitaire.Checked = (If(obj("Inegalitaire") IsNot Nothing, obj("Inegalitaire").ToString().ToUpper(), "") = "TRUE")
            RafraichirBDD()
            DtDepotCreateur.Rows.Clear()
            For Each ayant As AyantDroit In sacemData.AyantsDroit
                Dim nr As DataRow = DtDepotCreateur.NewRow()
                ' Chargement leger : seuls les champs BDO sont lus depuis le JSON
                ' L'identite est toujours enrichie depuis le XLSX via l'Id
                Dim idJson As String = ayant.BDO.Id
                nr("Id") = idJson
                nr("Role") = ayant.BDO.Role
                nr("Lettrage") = ayant.BDO.Lettrage
                nr("PH") = ayant.BDO.PH
                nr("Managelic") = ayant.BDO.Managelic
                nr("Managesub") = ayant.BDO.Managesub
                nr("Signataire") = ayant.BDO.Signataire
                ' Type : depuis JSON si verbeux, sinon depuis préfixe Id
                Dim typeJson As String = If(ayant.Identite IsNot Nothing, ayant.Identite.Type, "")
                If String.IsNullOrEmpty(typeJson) Then
                    typeJson = If(idJson.Trim().ToUpper().StartsWith("M"), "Moral", "Physique")
                End If
                nr("Type") = typeJson

                ' JSON verbeux = identité déjà présente → utiliser directement sans XLSX
                Dim isVerbeux As Boolean = ayant.Identite IsNot Nothing AndAlso
                    (Not String.IsNullOrEmpty(ayant.Identite.Nom) OrElse
                     Not String.IsNullOrEmpty(ayant.Identite.Designation))

                If isVerbeux Then
                    nr("Nom") = If(ayant.Identite.Nom, "")
                    nr("Prenom") = If(ayant.Identite.Prenom, "")
                    nr("Pseudonyme") = If(ayant.Identite.Pseudonyme, "")
                    nr("Designation") = If(ayant.Identite.Designation, "")
                    nr("Genre") = If(ayant.Identite.Genre, "")
                    nr("Nele") = If(ayant.Identite.Nele, "")
                    nr("Nea") = If(ayant.Identite.Nea, "")
                    nr("SocieteGestion") = If(ayant.Identite.SocieteGestion, "")
                    nr("FormeJuridique") = If(ayant.Identite.FormeJuridique, "")
                    nr("Capital") = If(ayant.Identite.Capital, "")
                    nr("RCS") = If(ayant.Identite.RCS, "")
                    nr("Siren") = If(ayant.Identite.Siren, "")
                    nr("GenreRepresentant") = If(ayant.Identite.GenreRepresentant, "")
                    nr("PrenomRepresentant") = If(ayant.Identite.PrenomRepresentant, "")
                    nr("NomRepresentant") = If(ayant.Identite.NomRepresentant, "")
                    nr("FonctionRepresentant") = If(ayant.Identite.FonctionRepresentant, "")
                    nr("NumVoie") = If(ayant.Adresse IsNot Nothing, If(ayant.Adresse.NumVoie, ""), "")
                    nr("TypeVoie") = If(ayant.Adresse IsNot Nothing, If(ayant.Adresse.TypeVoie, ""), "")
                    nr("NomVoie") = If(ayant.Adresse IsNot Nothing, If(ayant.Adresse.NomVoie, ""), "")
                    nr("CP") = If(ayant.Adresse IsNot Nothing, If(ayant.Adresse.CP, ""), "")
                    nr("Ville") = If(ayant.Adresse IsNot Nothing, If(ayant.Adresse.Ville, ""), "")
                    nr("Pays") = If(ayant.Adresse IsNot Nothing, If(ayant.Adresse.Pays, ""), "")
                    nr("Mail") = If(ayant.Contact IsNot Nothing, If(ayant.Contact.Mail, ""), "")
                    nr("Tel") = If(ayant.Contact IsNot Nothing, If(ayant.Contact.Tel, ""), "")
                    nr("COAD_IPI") = If(ayant.BDO.COAD_IPI, "")
                    nr("_Orphelin") = False
                ElseIf typeJson = "Physique" AndAlso DtPersonPhy IsNot Nothing Then
                    ' Format léger Physique : enrichissement depuis XLSX
                    Dim phyRow As DataRow = Nothing
                    If Not String.IsNullOrEmpty(idJson) AndAlso DtPersonPhy.Columns.Contains("Id") Then
                        phyRow = DtPersonPhy.AsEnumerable().FirstOrDefault(
                            Function(r) SafeStr(r, "Id").Trim().ToUpper() = idJson.Trim().ToUpper())
                    End If
                    nr("_Orphelin") = (phyRow Is Nothing)
                    If phyRow IsNot Nothing Then
                        nr("Pseudonyme") = SafeStr(phyRow, "Pseudonyme")
                        nr("Nom") = SafeStr(phyRow, "Nom")
                        nr("Prenom") = SafeStr(phyRow, "Prenom")
                        nr("Genre") = SafeStr(phyRow, "Genre")
                        nr("Nele") = SafeStr(phyRow, "Date de naissance")
                        nr("Nea") = SafeStr(phyRow, "Lieu de naissance")
                        nr("NumVoie") = SafeStr(phyRow, "Num de voie")
                        nr("TypeVoie") = SafeStr(phyRow, "Type de voie")
                        nr("NomVoie") = SafeStr(phyRow, "Nom de voie")
                        nr("CP") = SafeStr(phyRow, "CP")
                        nr("Ville") = SafeStr(phyRow, "Ville")
                        nr("Mail") = SafeStr(phyRow, "Mail")
                        nr("Tel") = SafeStr(phyRow, "Tel")
                        nr("COAD_IPI") = GetCOADIPI(phyRow)
                        Dim desigBuilt As String = BuildDesignation(phyRow)
                        If nr("Role").ToString().Trim().ToUpper() = "AEC" Then desigBuilt &= " (EAC)"
                        nr("Designation") = desigBuilt
                    End If
                ElseIf typeJson = "Moral" AndAlso DtPersonMor IsNot Nothing Then
                    ' Format léger Moral : enrichissement depuis XLSX
                    Dim morRow As DataRow = Nothing
                    If Not String.IsNullOrEmpty(idJson) AndAlso DtPersonMor.Columns.Contains("Id") Then
                        morRow = DtPersonMor.AsEnumerable().FirstOrDefault(
                            Function(r) SafeStr(r, "Id").Trim().ToUpper() = idJson.Trim().ToUpper())
                    End If
                    nr("_Orphelin") = (morRow Is Nothing)
                    If morRow IsNot Nothing Then
                        nr("FormeJuridique") = SafeStr(morRow, "Forme Juridique")
                        nr("Capital") = SafeStr(morRow, "Capital")
                        nr("RCS") = SafeStr(morRow, "RCS")
                        nr("Siren") = SafeStr(morRow, "Siren")
                        nr("PrenomRepresentant") = SafeStr(morRow, "Prenom representant")
                        nr("NomRepresentant") = SafeStr(morRow, "Nom representant")
                        nr("FonctionRepresentant") = SafeStr(morRow, "Fonction representant")
                        nr("NumVoie") = SafeStr(morRow, "Num de voie")
                        nr("TypeVoie") = SafeStr(morRow, "Type de voie")
                        nr("NomVoie") = SafeStr(morRow, "Nom de voie")
                        nr("CP") = SafeStr(morRow, "CP")
                        nr("Ville") = SafeStr(morRow, "Ville")
                        nr("Mail") = SafeStr(morRow, "Mail")
                        nr("Tel") = SafeStr(morRow, "Tel")
                        nr("COAD_IPI") = GetCOADIPI(morRow)
                        nr("Designation") = SafeStr(morRow, "Designation")
                    End If
                End If
                DtDepotCreateur.Rows.Add(nr)
            Next
            dgv.Refresh()
            AppelerMoteur()
            lblStatut.Text = "JSON charge : " & Path.GetFileName(filePath) & " — donnees enrichies depuis BDD"
            RefreshDeclarationFormat()
        Catch ex As Exception
            MessageBox.Show("Erreur chargement JSON : " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' GRILLE — ÉVÉNEMENTS
    ' ─────────────────────────────────────────────────────────────
    Private Sub Dgv_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs)
        If dgv.Columns(e.ColumnIndex).Name = "Role" Then
            Dim val As String = e.FormattedValue.ToString().Trim().ToUpper()
            Dim valid As String() = {"A", "C", "AR", "AD", "E", "AEC"}
            If Not valid.Contains(val) Then
                e.Cancel = True
                MessageBox.Show("Roles valides : A, C, AR, AD, E, AEC", "Valeur invalide", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub Dgv_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        Dim editable As String() = {"Role", "Lettrage", "PH", "SocieteGestion", "Signataire", "Managelic", "Managesub", "COAD_IPI"}
        If Not editable.Contains(dgv.Columns(e.ColumnIndex).Name) Then
            e.Cancel = True
            Return
        End If
        If dgv.Columns(e.ColumnIndex).Name = "PH" AndAlso e.RowIndex >= 0 Then
            Dim row As DataRow = DtDepotCreateur.Rows(e.RowIndex)
            Dim role As String = row("Role").ToString().Trim().ToUpper()
            Dim lettrage As String = row("Lettrage").ToString().Trim()
            If role = "A" OrElse role = "C" Then
                _totalAvantPH = DtDepotCreateur.AsEnumerable().Where(Function(r) r("Role").ToString().Trim().ToUpper() = role).Sum(Function(r) MoteurRepartition.ParsePH(r))
            ElseIf MoteurRepartition.IsEditeur(role) Then
                _totalAvantPH = DtDepotCreateur.AsEnumerable().Where(Function(r) MoteurRepartition.IsEditeur(r("Role").ToString()) AndAlso
                                      r("Lettrage").ToString().Trim() = lettrage).Sum(Function(r) MoteurRepartition.ParsePH(r))
            Else
                _totalAvantPH = 0
            End If
        End If
    End Sub

    Private Sub Dgv_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        e.ThrowException = False
        Dim colName As String = dgv.Columns(e.ColumnIndex).Name
        If colName = "Managelic" OrElse colName = "Managesub" Then
            Dim cbCol As DataGridViewComboBoxColumn = TryCast(dgv.Columns(e.ColumnIndex), DataGridViewComboBoxColumn)
            If cbCol IsNot Nothing AndAlso e.RowIndex >= 0 AndAlso e.RowIndex < DtDepotCreateur.Rows.Count Then
                Dim rowVal As String = DtDepotCreateur.Rows(e.RowIndex)(colName).ToString()
                If Not String.IsNullOrEmpty(rowVal) AndAlso Not cbCol.Items.Contains(rowVal) Then
                    cbCol.Items.Add(rowVal)
                End If
            End If
        End If
    End Sub

    Private Sub Dgv_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Dim hit = dgv.HitTest(e.X, e.Y)
            If hit.RowIndex >= 0 Then
                dgv.ClearSelection()
                dgv.Rows(hit.RowIndex).Selected = True
                cmsGrille.Show(dgv, e.Location)
            End If
        End If
    End Sub

    Private Sub MnuSupprimer_Click(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then Return
        Dim selRow As DataGridViewRow = dgv.SelectedRows(0)
        Dim lettrage As String = selRow.Cells("Lettrage").Value.ToString()
        Dim role As String = selRow.Cells("Role").Value.ToString()
        If role = "A" OrElse role = "C" OrElse role = "AD" OrElse role = "AR" Then
            For i As Integer = DtDepotCreateur.Rows.Count - 1 To 0 Step -1
                Dim r As DataRow = DtDepotCreateur.Rows(i)
                If r("Lettrage").ToString() = lettrage AndAlso MoteurRepartition.IsEditeur(r("Role").ToString()) Then
                    DtDepotCreateur.Rows.Remove(r)
                End If
            Next
        End If
        dgv.Rows.Remove(selRow)
        UpdateLettrages()
        RefreshDeclarationFormat()
        dgv.Refresh()
    End Sub

    Private Sub Dgv_MouseDown(sender As Object, e As MouseEventArgs)
        Dim hit = dgv.HitTest(e.X, e.Y)
        If hit.RowIndex >= 0 Then _dragRowIndex = hit.RowIndex
    End Sub

    Private Sub Dgv_MouseMove(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left AndAlso _dragRowIndex >= 0 Then
            dgv.DoDragDrop(_dragRowIndex, DragDropEffects.Move)
        End If
    End Sub

    Private Sub Dgv_DragOver(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.Move
    End Sub

    Private Sub Dgv_DragDrop(sender As Object, e As DragEventArgs)
        If Not e.Data.GetDataPresent(GetType(Integer)) Then Return
        Dim sourceIndex As Integer = CInt(e.Data.GetData(GetType(Integer)))
        Dim clientPoint As Point = dgv.PointToClient(New Point(e.X, e.Y))
        Dim hit = dgv.HitTest(clientPoint.X, clientPoint.Y)
        Dim destIndex As Integer = If(hit.RowIndex >= 0, hit.RowIndex, dgv.Rows.Count - 1)
        If sourceIndex = destIndex Then Return
        Dim sourceRow As DataRow = DtDepotCreateur.Rows(sourceIndex)
        Dim newRow As DataRow = DtDepotCreateur.NewRow()
        newRow.ItemArray = sourceRow.ItemArray.Clone()
        DtDepotCreateur.Rows.RemoveAt(sourceIndex)
        If destIndex > DtDepotCreateur.Rows.Count Then
            DtDepotCreateur.Rows.Add(newRow)
        Else
            DtDepotCreateur.Rows.InsertAt(newRow, destIndex)
        End If
        dgv.ClearSelection()
        If destIndex < dgv.Rows.Count Then dgv.Rows(destIndex).Selected = True
        ApplyRowColors()
        RefreshDeclarationFormat()
        _dragRowIndex = -1
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' RECHERCHE LIVE
    ' ─────────────────────────────────────────────────────────────
    Private Sub TxtRecherche_GotFocus(sender As Object, e As EventArgs)
        If _watermark Then
            txtRecherche.Text = ""
            txtRecherche.ForeColor = Color.Black
            _watermark = False
        End If
    End Sub

    Private Sub TxtRecherche_LostFocus(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtRecherche.Text) Then
            txtRecherche.Text = "Nom, prenom, pseudonyme..."
            txtRecherche.ForeColor = Color.Gray
            _watermark = True
            lstResultats.Visible = False
        End If
    End Sub

    Private Sub TxtRecherche_Changed(sender As Object, e As EventArgs)
        If _watermark Then Return
        Dim terme As String = txtRecherche.Text.Trim().ToUpper()
        lstResultats.Items.Clear()
        If terme.Length = 0 Then lstResultats.Visible = False : Return
        If DtPersonPhy IsNot Nothing Then
            For Each r As DataRow In DtPersonPhy.Rows
                If SafeStr(r, "Nom").ToUpper().Contains(terme) OrElse
                   SafeStr(r, "Prenom").ToUpper().Contains(terme) OrElse
                   SafeStr(r, "Pseudonyme").ToUpper().Contains(terme) Then
                    lstResultats.Items.Add($"{SafeStr(r, "Pseudonyme")} / {SafeStr(r, "Nom")} / {SafeStr(r, "Prenom")}")
                End If
            Next
        End If
        If DtPersonMor IsNot Nothing Then
            For Each r As DataRow In DtPersonMor.Rows
                If SafeStr(r, "Designation").ToUpper().Contains(terme) Then
                    lstResultats.Items.Add("(Morale) " & SafeStr(r, "Designation"))
                End If
            Next
        End If
        lstResultats.Visible = lstResultats.Items.Count > 0
    End Sub

    Private Sub LstResultats_DoubleClick(sender As Object, e As EventArgs)
        If lstResultats.SelectedItem Is Nothing Then Return
        Dim sel As String = lstResultats.SelectedItem.ToString()
        RafraichirBDD()
        If sel.StartsWith("(Morale) ") Then
            AjouterPersonneMorale(sel.Substring(9).Trim())
        Else
            AjouterPersonnePhysique(sel)
        End If
        AppelerMoteur()
        UpdateLettrages()
        ApplyRowColors()
        RefreshDeclarationFormat()
        lstResultats.Visible = False
        txtRecherche.Text = ""
        _watermark = False
        dgv.Refresh()
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' DOUBLE-CLIC GRILLE
    ' ─────────────────────────────────────────────────────────────
    Private Sub Dgv_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim row As DataRow = DtDepotCreateur.Rows(e.RowIndex)
        If row("Type").ToString().Trim() = "Physique" Then
            OuvrirFichePhysique(row)
        ElseIf row("Type").ToString().Trim() = "Moral" Then
            OuvrirFicheMorale(row)
        End If
    End Sub

    Private Sub OuvrirFichePhysique(row As DataRow)
        Dim srcRow As DataRow = Nothing
        If DtPersonPhy IsNot Nothing Then
            Dim idRow As String = row("Id").ToString().Trim()
            If Not String.IsNullOrEmpty(idRow) Then
                srcRow = DtPersonPhy.AsEnumerable().FirstOrDefault(
                    Function(r) SafeStr(r, "Id").Trim().ToUpper() = idRow.ToUpper())
            End If
            If srcRow Is Nothing Then
                Dim nom As String = row("Nom").ToString().Trim()
                Dim prenom As String = row("Prenom").ToString().Trim()
                Dim pseudo As String = row("Pseudonyme").ToString().Trim()
                For Each r As DataRow In DtPersonPhy.Rows
                    If (Not String.IsNullOrEmpty(nom) AndAlso SafeStr(r, "Nom").Trim() = nom AndAlso SafeStr(r, "Prenom").Trim() = prenom) OrElse
                       (Not String.IsNullOrEmpty(pseudo) AndAlso SafeStr(r, "Pseudonyme").Trim() = pseudo) Then
                        srcRow = r : Exit For
                    End If
                Next
            End If
        End If
        Dim cols() As String = PersonnesForm.ColsPhy
        Dim vals(cols.Length - 1) As String
        For i As Integer = 0 To cols.Length - 1
            Dim col As String = cols(i)
            If srcRow IsNot Nothing Then
                Try : vals(i) = srcRow(col).ToString() : Catch : vals(i) = "" : End Try
            End If
            If String.IsNullOrEmpty(vals(i)) Then
                Select Case col
                    Case "Pseudonyme" : vals(i) = row("Pseudonyme").ToString()
                    Case "Nom" : vals(i) = row("Nom").ToString()
                    Case "Prenom" : vals(i) = row("Prenom").ToString()
                    Case "Genre" : vals(i) = row("Genre").ToString()
                    Case "Role" : vals(i) = row("Role").ToString()
                    Case "COAD"
                        Dim ci As String = row("COAD_IPI").ToString()
                        If ci.StartsWith("COAD : ") Then vals(i) = ci.Substring(7).Trim()
                    Case "IPI"
                        Dim ci As String = row("COAD_IPI").ToString()
                        If ci.StartsWith("IPI : ") Then vals(i) = ci.Substring(6).Trim()
                    Case "Num de voie" : vals(i) = row("NumVoie").ToString()
                    Case "Type de voie" : vals(i) = row("TypeVoie").ToString()
                    Case "Nom de voie" : vals(i) = row("NomVoie").ToString()
                    Case "CP" : vals(i) = row("CP").ToString()
                    Case "Ville" : vals(i) = row("Ville").ToString()
                    Case "Mail" : vals(i) = row("Mail").ToString()
                    Case "Tel" : vals(i) = row("Tel").ToString()
                End Select
            End If
        Next
        Using f As New FichePersonneForm(cols, vals, "Modifier Personne Physique")
            If f.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim res() As String = f.GetValues()
            If srcRow IsNot Nothing Then
                For i As Integer = 0 To cols.Length - 1
                    Try : srcRow(cols(i)) = res(i) : Catch : End Try
                Next
            End If
            Dim iId     As Integer = Array.IndexOf(cols, "Id")
            Dim iPseudo As Integer = Array.IndexOf(cols, "Pseudonyme")
            Dim iNom    As Integer = Array.IndexOf(cols, "Nom")
            Dim iPrenom As Integer = Array.IndexOf(cols, "Prenom")
            Dim iGenre  As Integer = Array.IndexOf(cols, "Genre")
            Dim iRole   As Integer = Array.IndexOf(cols, "Role")
            Dim iCOAD   As Integer = Array.IndexOf(cols, "COAD")
            Dim iIPI    As Integer = Array.IndexOf(cols, "IPI")
            Dim iSG     As Integer = Array.IndexOf(cols, "SocieteGestion")
            Dim iNumV   As Integer = Array.IndexOf(cols, "Num de voie")
            Dim iTypV   As Integer = Array.IndexOf(cols, "Type de voie")
            Dim iNomV   As Integer = Array.IndexOf(cols, "Nom de voie")
            Dim iCP     As Integer = Array.IndexOf(cols, "CP")
            Dim iVille  As Integer = Array.IndexOf(cols, "Ville")
            Dim iMail   As Integer = Array.IndexOf(cols, "Mail")
            Dim iTel    As Integer = Array.IndexOf(cols, "Tel")
            If iPseudo >= 0 Then row("Pseudonyme") = res(iPseudo)
            If iNom    >= 0 Then row("Nom")        = res(iNom)
            If iPrenom >= 0 Then row("Prenom")     = res(iPrenom)
            If iGenre  >= 0 Then row("Genre")      = res(iGenre)
            If iRole   >= 0 AndAlso row("Role").ToString().Trim().ToUpper() <> "AEC" Then row("Role") = res(iRole)
            Dim coad As String = If(iCOAD >= 0, res(iCOAD).Trim(), "")
            Dim ipi  As String = If(iIPI  >= 0, res(iIPI).Trim(), "")
            If Not String.IsNullOrEmpty(coad) Then
                row("COAD_IPI") = "COAD : " & coad
            ElseIf Not String.IsNullOrEmpty(ipi) Then
                row("COAD_IPI") = "IPI : " & ipi
            End If
            If iSG   >= 0 Then row("SocieteGestion") = res(iSG)
            If iNumV >= 0 Then row("NumVoie")        = res(iNumV)
            If iTypV >= 0 Then row("TypeVoie")       = res(iTypV)
            If iNomV >= 0 Then row("NomVoie")        = res(iNomV)
            If iCP   >= 0 Then row("CP")             = res(iCP)
            If iVille >= 0 Then row("Ville")         = res(iVille)
            If iMail >= 0 Then row("Mail")           = res(iMail)
            If iTel  >= 0 Then row("Tel")            = res(iTel)
            Dim nom2    As String = If(iNom    >= 0, res(iNom).Trim(), "")
            Dim prenom2 As String = If(iPrenom >= 0, res(iPrenom).Trim(), "")
            Dim pseudo2 As String = If(iPseudo >= 0, res(iPseudo).Trim(), "")
            Dim desigBase As String = If(Not String.IsNullOrEmpty(pseudo2),
                                         nom2 & " " & prenom2 & " / " & pseudo2,
                                         (nom2 & " " & prenom2).Trim())
            If row("Role").ToString().Trim().ToUpper() = "AEC" Then desigBase &= " (EAC)"
            row("Designation") = desigBase
            SauvegarderXlsxSilencieux()
            dgv.Refresh()
            ApplyRowColors()
            lblStatut.Text = "Personne physique mise a jour."
        End Using
    End Sub

    Private Sub OuvrirFicheMorale(row As DataRow)
        Dim srcRow As DataRow = Nothing
        If DtPersonMor IsNot Nothing Then
            Dim idRow As String = row("Id").ToString().Trim()
            If Not String.IsNullOrEmpty(idRow) Then
                srcRow = DtPersonMor.AsEnumerable().FirstOrDefault(
                    Function(r) SafeStr(r, "Id").Trim().ToUpper() = idRow.ToUpper())
            End If
            If srcRow Is Nothing Then
                Dim desig As String = row("Designation").ToString().Trim().ToUpper()
                For Each r As DataRow In DtPersonMor.Rows
                    If SafeStr(r, "Designation").Trim().ToUpper() = desig Then srcRow = r : Exit For
                Next
            End If
        End If
        Dim cols() As String = PersonnesForm.ColsMor
        Dim vals(cols.Length - 1) As String
        For i As Integer = 0 To cols.Length - 1
            Dim col As String = cols(i)
            If srcRow IsNot Nothing Then
                Try : vals(i) = srcRow(col).ToString() : Catch : vals(i) = "" : End Try
            End If
            If String.IsNullOrEmpty(vals(i)) Then
                Select Case col
                    Case "Designation" : vals(i) = row("Designation").ToString()
                    Case "COAD"
                        Dim ci As String = row("COAD_IPI").ToString()
                        If ci.StartsWith("COAD : ") Then vals(i) = ci.Substring(7).Trim()
                    Case "IPI"
                        Dim ci As String = row("COAD_IPI").ToString()
                        If ci.StartsWith("IPI : ") Then vals(i) = ci.Substring(6).Trim()
                    Case "Forme Juridique" : vals(i) = row("FormeJuridique").ToString()
                    Case "Capital" : vals(i) = row("Capital").ToString()
                    Case "RCS" : vals(i) = row("RCS").ToString()
                    Case "Siren" : vals(i) = row("Siren").ToString()
                    Case "Num de voie" : vals(i) = row("NumVoie").ToString()
                    Case "Type de voie" : vals(i) = row("TypeVoie").ToString()
                    Case "Nom de voie" : vals(i) = row("NomVoie").ToString()
                    Case "CP" : vals(i) = row("CP").ToString()
                    Case "Ville" : vals(i) = row("Ville").ToString()
                    Case "Prenom representant" : vals(i) = row("PrenomRepresentant").ToString()
                    Case "Nom representant" : vals(i) = row("NomRepresentant").ToString()
                    Case "Fonction representant" : vals(i) = row("FonctionRepresentant").ToString()
                    Case "Mail" : vals(i) = row("Mail").ToString()
                    Case "Tel" : vals(i) = row("Tel").ToString()
                End Select
            End If
        Next
        Using f As New FichePersonneForm(cols, vals, "Modifier Personne Morale")
            If f.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim res() As String = f.GetValues()
            If srcRow IsNot Nothing Then
                For i As Integer = 0 To cols.Length - 1
                    Try : srcRow(cols(i)) = res(i) : Catch : End Try
                Next
            End If
            Dim iDesig  As Integer = Array.IndexOf(cols, "Designation")
            Dim iCOAD   As Integer = Array.IndexOf(cols, "COAD")
            Dim iIPI    As Integer = Array.IndexOf(cols, "IPI")
            Dim iFJ     As Integer = Array.IndexOf(cols, "Forme Juridique")
            Dim iCap    As Integer = Array.IndexOf(cols, "Capital")
            Dim iRCS    As Integer = Array.IndexOf(cols, "RCS")
            Dim iSiren  As Integer = Array.IndexOf(cols, "Siren")
            Dim iSG     As Integer = Array.IndexOf(cols, "SocieteGestion")
            Dim iNumV   As Integer = Array.IndexOf(cols, "Num de voie")
            Dim iTypV   As Integer = Array.IndexOf(cols, "Type de voie")
            Dim iNomV   As Integer = Array.IndexOf(cols, "Nom de voie")
            Dim iCP     As Integer = Array.IndexOf(cols, "CP")
            Dim iVille  As Integer = Array.IndexOf(cols, "Ville")
            Dim iPrenR  As Integer = Array.IndexOf(cols, "Prenom representant")
            Dim iNomR   As Integer = Array.IndexOf(cols, "Nom representant")
            Dim iFoncR  As Integer = Array.IndexOf(cols, "Fonction representant")
            Dim iMail   As Integer = Array.IndexOf(cols, "Mail")
            Dim iTel    As Integer = Array.IndexOf(cols, "Tel")
            If iDesig >= 0 Then row("Designation") = res(iDesig)
            Dim coad As String = If(iCOAD >= 0, res(iCOAD).Trim(), "")
            Dim ipi  As String = If(iIPI  >= 0, res(iIPI).Trim(), "")
            If Not String.IsNullOrEmpty(coad) Then
                row("COAD_IPI") = "COAD : " & coad
            ElseIf Not String.IsNullOrEmpty(ipi) Then
                row("COAD_IPI") = "IPI : " & ipi
            End If
            If iFJ    >= 0 Then row("FormeJuridique")       = res(iFJ)
            If iCap   >= 0 Then row("Capital")              = res(iCap)
            If iRCS   >= 0 Then row("RCS")                  = res(iRCS)
            If iSiren >= 0 Then row("Siren")                = res(iSiren)
            If iSG    >= 0 Then row("SocieteGestion")       = res(iSG)
            If iNumV  >= 0 Then row("NumVoie")              = res(iNumV)
            If iTypV  >= 0 Then row("TypeVoie")             = res(iTypV)
            If iNomV  >= 0 Then row("NomVoie")              = res(iNomV)
            If iCP    >= 0 Then row("CP")                   = res(iCP)
            If iVille >= 0 Then row("Ville")                = res(iVille)
            If iPrenR >= 0 Then row("PrenomRepresentant")   = res(iPrenR)
            If iNomR  >= 0 Then row("NomRepresentant")      = res(iNomR)
            If iFoncR >= 0 Then row("FonctionRepresentant") = res(iFoncR)
            If iMail  >= 0 Then row("Mail")                 = res(iMail)
            If iTel   >= 0 Then row("Tel")                  = res(iTel)
            SauvegarderXlsxSilencieux()
            dgv.Refresh()
            ApplyRowColors()
            lblStatut.Text = "Personne morale mise a jour."
        End Using
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' GRILLE — PH / INÉGALITAIRE
    ' ─────────────────────────────────────────────────────────────
    Private Sub Dgv_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If dgv.IsCurrentCellDirty Then
            If dgv.CurrentCell IsNot Nothing AndAlso
               dgv.Columns(dgv.CurrentCell.ColumnIndex).Name = "PH" Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            End If
        End If
    End Sub

    Private Sub Dgv_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If dgv.Columns(e.ColumnIndex).Name = "PH" Then
            Try
                Dim row As DataRow = DtDepotCreateur.Rows(e.RowIndex)
                Dim params As MoteurRepartition.ParamsOeuvre = GetParamsOeuvre()
                MoteurRepartition.RééquilibrerCategorie(DtDepotCreateur, row, _totalAvantPH)
                MoteurRepartition.Calculer(DtDepotCreateur, params)
            Catch ex As Exception
                lblStatut.Text = "Erreur recalcul PH : " & ex.Message
            End Try
            ApplyRowColors()
            dgv.Refresh()
        End If
    End Sub

    Private Sub CbInegalitaire_CheckedChanged(sender As Object, e As EventArgs)
        AppelerMoteur()
        ApplyRowColors()
        dgv.Refresh()
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' HELPERS
    ' ─────────────────────────────────────────────────────────────
    Private Function GetNextLetter() As String
        Dim used As New HashSet(Of String)(
            DtDepotCreateur.AsEnumerable().Select(Function(r) r("Lettrage").ToString()))
        For Each c As Char In "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If Not used.Contains(c.ToString()) Then Return c.ToString()
        Next
        Return ""
    End Function

    Private Sub UpdateLettrages()
        For Each row As DataRow In DtDepotCreateur.Rows
            Dim role As String = row("Role").ToString()
            If (role = "A" OrElse role = "C" OrElse role = "AD" OrElse role = "AR") AndAlso
               String.IsNullOrEmpty(row("Lettrage").ToString()) Then
                row("Lettrage") = GetNextLetter()
            End If
        Next
    End Sub

    Private Sub ApplyRowColors()
        For Each row As DataGridViewRow In dgv.Rows
            Dim dataRow As DataRow = DirectCast(row.DataBoundItem, DataRowView).Row
            Dim mail As String = dataRow("Mail").ToString().Trim()
            Dim desig As String = dataRow("Designation").ToString().Trim()
            Dim orphelin As Boolean = False
            If dataRow.Table.Columns.Contains("_Orphelin") Then
                Boolean.TryParse(dataRow("_Orphelin").ToString(), orphelin)
            End If
            If orphelin Then
                row.DefaultCellStyle.BackColor = Color.DarkOrange
                row.DefaultCellStyle.ForeColor = Color.White
                row.DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Italic)
            ElseIf String.IsNullOrEmpty(desig) Then
                row.DefaultCellStyle.BackColor = Color.Black
                row.DefaultCellStyle.ForeColor = Color.White
                row.DefaultCellStyle.Font = dgv.Font
            ElseIf String.IsNullOrEmpty(mail) Then
                row.DefaultCellStyle.BackColor = Color.OrangeRed
                row.DefaultCellStyle.ForeColor = Color.White
                row.DefaultCellStyle.Font = dgv.Font
            ElseIf Not String.IsNullOrEmpty(dataRow("CP").ToString()) Then
                row.DefaultCellStyle.BackColor = Color.LightGreen
                row.DefaultCellStyle.ForeColor = Color.Black
                row.DefaultCellStyle.Font = dgv.Font
            Else
                row.DefaultCellStyle.BackColor = Color.LightYellow
                row.DefaultCellStyle.ForeColor = Color.Black
                row.DefaultCellStyle.Font = dgv.Font
            End If
        Next
    End Sub

    Private Function SafeStr(row As DataRow, colName As String, Optional defVal As String = "") As String
        If row.Table.Columns.Contains(colName) AndAlso Not IsDBNull(row(colName)) Then
            Return row(colName).ToString()
        End If
        Dim colNorm As String = NormalizeCol(colName)
        For Each col As DataColumn In row.Table.Columns
            If NormalizeCol(col.ColumnName) = colNorm Then
                If Not IsDBNull(row(col)) Then Return row(col).ToString()
                Return defVal
            End If
        Next
        Return defVal
    End Function

    Private Function NormalizeCol(s As String) As String
        s = s.ToLower().Trim()
        s = s.Replace("é", "e").Replace("è", "e").Replace("ê", "e")
        s = s.Replace("à", "a").Replace("â", "a")
        s = s.Replace("ô", "o").Replace("î", "i").Replace("û", "u")
        s = s.Replace("ç", "c").Replace("°", "").Replace("ô", "o")
        Return s
    End Function

    Private Function GetCOADIPI(row As DataRow) As String
        Dim coad As String = SafeStr(row, "COAD")
        Dim ipi As String = SafeStr(row, "IPI")
        If Not String.IsNullOrEmpty(coad) Then Return "COAD : " & coad
        If Not String.IsNullOrEmpty(ipi) Then Return "IPI : " & ipi
        Return ""
    End Function

    Private Function BuildDesignation(row As DataRow) As String
        Dim nom As String = SafeStr(row, "Nom")
        Dim prenom As String = SafeStr(row, "Prenom")
        Dim pseudo As String = SafeStr(row, "Pseudonyme")
        Return $"{nom} {prenom}".Trim() & If(String.IsNullOrEmpty(pseudo), "", $" / {pseudo}")
    End Function

    Private Sub CopierAdresseContact(dest As DataRow, source As DataRow)
        dest("NumVoie") = SafeStr(source, "Num de voie")
        dest("TypeVoie") = SafeStr(source, "Type de voie")
        dest("NomVoie") = SafeStr(source, "Nom de voie")
        dest("CP") = SafeStr(source, "CP")
        dest("Ville") = SafeStr(source, "Ville")
        dest("Pays") = SafeStr(source, "Pays", "FRANCE")
        dest("Mail") = SafeStr(source, "Mail")
        dest("Tel") = SafeStr(source, "Tel")
    End Sub

    Private Sub CopierInfosMorale(dest As DataRow, source As DataRow)
        dest("FormeJuridique") = SafeStr(source, "Forme Juridique")
        dest("Capital") = SafeStr(source, "Capital")
        dest("RCS") = SafeStr(source, "RCS")
        dest("Siren") = SafeStr(source, "Siren")
        dest("GenreRepresentant") = SafeStr(source, "Genre representant")
        dest("PrenomRepresentant") = SafeStr(source, "Prenom representant")
        dest("NomRepresentant") = SafeStr(source, "Nom representant")
        dest("FonctionRepresentant") = SafeStr(source, "Fonction representant")
        Dim coad As String = SafeStr(source, "COAD")
        Dim ipi As String = SafeStr(source, "IPI")
        If Not String.IsNullOrEmpty(coad) Then
            dest("COAD_IPI") = "COAD : " & coad
        ElseIf Not String.IsNullOrEmpty(ipi) Then
            dest("COAD_IPI") = "IPI : " & ipi
        End If
    End Sub

    Private Function AjusterRole(role As String, genre As String) As String
        Dim litteraires As String() = {"Billet d'humeur", "Texte", "Texte de presentation",
                                       "Texte de sketch", "Poeme", "Chronique"}
        Dim musicaux As String() = {"Instrumental", "Musique illustrative"}
        If litteraires.Contains(genre) AndAlso role = "C" Then Return "A"
        If musicaux.Contains(genre) AndAlso role = "A" Then Return "C"
        Return role
    End Function

    Private Function RolesPossibles(genre As String) As List(Of String)
        Dim litteraires As String() = {"Billet d'humeur", "Texte", "Texte de presentation",
                                       "Texte de sketch", "Poeme", "Chronique"}
        Dim musicaux As String() = {"Instrumental", "Musique illustrative"}
        If litteraires.Contains(genre) Then Return New List(Of String) From {"A", "AD"}
        If musicaux.Contains(genre) Then Return New List(Of String) From {"C", "AR"}
        Return New List(Of String) From {"A", "C", "AR", "AD"}
    End Function

    Private Function CleanFileName(name As String) As String
        For Each c As Char In Path.GetInvalidFileNameChars()
            name = name.Replace(c, CChar("_"))
        Next
        Return name
    End Function

End Class

' ─────────────────────────────────────────────────────────────────────────────
' FORMULAIRE SÉLECTION CRÉATEURS
' ─────────────────────────────────────────────────────────────────────────────
Public Class FormCreateurs
    Inherits Form

    Public ReadOnly Property SelectedCreateurs As List(Of String)
        Get
            Return clbCreateurs.CheckedItems.Cast(Of String).ToList()
        End Get
    End Property

    Private clbCreateurs As CheckedListBox

    Public Sub New(createurs As List(Of String))
        Me.Text = "Selectionner les createurs associes"
        Me.Size = New Size(350, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        clbCreateurs = New CheckedListBox()
        clbCreateurs.Location = New Point(10, 10)
        clbCreateurs.Size = New Size(310, 200)
        clbCreateurs.Items.AddRange(createurs.ToArray())
        Me.Controls.Add(clbCreateurs)
        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Location = New Point(130, 225)
        btnOK.Size = New Size(80, 28)
        btnOK.DialogResult = DialogResult.OK
        Me.Controls.Add(btnOK)
        Me.AcceptButton = btnOK
    End Sub
End Class

' ─────────────────────────────────────────────────────────────────────────────
' FORMULAIRE SÉLECTION RÔLE
' ─────────────────────────────────────────────────────────────────────────────
Public Class FormRoles
    Inherits Form

    Public ReadOnly Property SelectedRole As String
        Get
            Return If(lstRoles.SelectedItem IsNot Nothing, lstRoles.SelectedItem.ToString(), "")
        End Get
    End Property

    Private lstRoles As ListBox

    Public Sub New(roles As List(Of String))
        Me.Text = "Selectionner le role"
        Me.Size = New Size(250, 220)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        lstRoles = New ListBox()
        lstRoles.Location = New Point(10, 10)
        lstRoles.Size = New Size(210, 140)
        lstRoles.Items.AddRange(roles.ToArray())
        Me.Controls.Add(lstRoles)
        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Location = New Point(80, 160)
        btnOK.Size = New Size(80, 28)
        Me.Controls.Add(btnOK)
        Me.AcceptButton = btnOK
        AddHandler btnOK.Click, AddressOf BtnOK_Click
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        If lstRoles.SelectedItem IsNot Nothing Then
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            MessageBox.Show("Selectionnez un role.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
End Class
