Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq
Imports OfficeOpenXml

''' <summary>
''' Formulaire de création et d'édition des fichiers JSON SACEM.
''' Intègre la logique de visualseize/Form1 avec export au format SACEMJsonReader.
''' Nouveaux champs : ISWC, Declaration, Format, Faita, Faitle (œuvre)
'''                   SocieteGestion, Signataire, Managelic, Managesub (ayant droit)
''' </summary>
Public Class JsonEditorForm
    Inherits Form

    ' ─────────────────────────────────────────────────────────────
    ' DONNÉES
    ' ─────────────────────────────────────────────────────────────
    Private DtPersonPhy As DataTable
    Private DtPersonMor As DataTable
    Private DtDepotCreateur As DataTable


    ''' <summary>Chemin du JSON sauvegardé — renvoyé à MainForm après fermeture.</summary>
    Public Property SavedJsonPath As String

    ' ─────────────────────────────────────────────────────────────
    ' CONTRÔLES — ŒUVRE
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

    ' ─────────────────────────────────────────────────────────────
    ' CONTRÔLES — AYANTS DROIT
    ' ─────────────────────────────────────────────────────────────
    Private WithEvents cbPersonnes As ComboBox
    Private WithEvents btnAjouter As Button
    Private WithEvents dgv As DataGridView
    Private WithEvents btnCalculer As Button
    Private WithEvents btnSauvegarder As Button
    Private WithEvents btnChargerJson As Button
    Private WithEvents btnGererFiches As Button
    Private WithEvents lblStatut As Label

    Private WithEvents cmsGrille As ContextMenuStrip
    Private WithEvents mnuSupprimer As ToolStripMenuItem
    Private _dragRowIndex As Integer = -1

    ' ─────────────────────────────────────────────────────────────
    ' CONSTRUCTEUR
    ' ─────────────────────────────────────────────────────────────
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(existingJsonPath As String)
        InitializeComponent()
        If Not String.IsNullOrEmpty(existingJsonPath) AndAlso File.Exists(existingJsonPath) Then
            LoadJsonData(existingJsonPath)
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' INITIALISATION DES COMPOSANTS
    ' ─────────────────────────────────────────────────────────────
    Private Sub InitializeComponent()
        Me.Text = "Éditeur JSON SACEM"
        Me.Size = New Size(1160, 720)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimumSize = New Size(1000, 650)

        ' ── GroupBox ŒUVRE ──────────────────────────────────────
        Dim grpOeuvre As New GroupBox()
        grpOeuvre.Text = "Informations de l'œuvre"
        grpOeuvre.Location = New Point(10, 10)
        grpOeuvre.Size = New Size(1120, 205)
        grpOeuvre.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        Dim y1 As Integer = 22
        Dim y2 As Integer = 50
        Dim y3 As Integer = 78
        Dim y4 As Integer = 106
        Dim y5 As Integer = 134
        Dim y6 As Integer = 162

        ' Ligne 1 : Titre / Sous-titre
        grpOeuvre.Controls.Add(MkLabel("Titre :", 8, y1))
        txtTitre = MkTextBox(70, y1, 480)
        grpOeuvre.Controls.Add(txtTitre)
        grpOeuvre.Controls.Add(MkLabel("Sous-titre :", 565, y1))
        txtSousTitre = MkTextBox(648, y1, 460)
        grpOeuvre.Controls.Add(txtSousTitre)

        ' Ligne 2 : Interprète / ISWC
        grpOeuvre.Controls.Add(MkLabel("Interprète :", 8, y2))
        txtInterprete = MkTextBox(90, y2, 370)
        grpOeuvre.Controls.Add(txtInterprete)
        grpOeuvre.Controls.Add(MkLabel("ISWC :", 475, y2))
        txtISWC = MkTextBox(520, y2, 150)
        grpOeuvre.Controls.Add(txtISWC)
        grpOeuvre.Controls.Add(MkLabel("Durée :", 685, y2))
        txtDuree = MkTextBox(730, y2, 90)
        grpOeuvre.Controls.Add(txtDuree)
        grpOeuvre.Controls.Add(MkLabel("Genre :", 835, y2))
        cbGenre = New ComboBox()
        cbGenre.Location = New Point(880, y2)
        cbGenre.Size = New Size(228, 23)
        cbGenre.Items.AddRange(New Object() {"Chant", "Jazz", "Techno", "Musique de film",
            "Musique symphonique", "Instrumental", "Musique illustrative",
            "Billet d'humeur", "Texte", "Texte de présentation",
            "Texte de sketch", "Poeme", "Chronique"})
        cbGenre.Text = "Chant"
        grpOeuvre.Controls.Add(cbGenre)

        ' Ligne 3 : Date / Lieu / Territoire
        grpOeuvre.Controls.Add(MkLabel("Date exploit :", 8, y3))
        dtDate = New DateTimePicker()
        dtDate.Location = New Point(90, y3)
        dtDate.Size = New Size(110, 23)
        dtDate.Format = DateTimePickerFormat.Custom
        dtDate.CustomFormat = "dd/MM/yyyy"
        grpOeuvre.Controls.Add(dtDate)
        grpOeuvre.Controls.Add(MkLabel("Lieu :", 215, y3))
        cbLieu = New ComboBox()
        cbLieu.Location = New Point(250, y3)
        cbLieu.Size = New Size(430, 23)
        cbLieu.Items.Add("Deezer / Spotify / Amazon Music / Youtube/ Apple Music")
        grpOeuvre.Controls.Add(cbLieu)
        grpOeuvre.Controls.Add(MkLabel("Territoire :", 695, y3))
        cbTerritoire = New ComboBox()
        cbTerritoire.Location = New Point(760, y3)
        cbTerritoire.Size = New Size(348, 23)
        cbTerritoire.Items.Add("Monde")
        grpOeuvre.Controls.Add(cbTerritoire)

        ' Ligne 4 : Arrangement / Inégalitaire
        grpOeuvre.Controls.Add(MkLabel("Arrangement :", 8, y4))
        cbArrangement = New ComboBox()
        cbArrangement.Location = New Point(95, y4)
        cbArrangement.Size = New Size(340, 23)
        cbArrangement.Items.Add("Toutes")
        grpOeuvre.Controls.Add(cbArrangement)
        cbInegalitaire = New CheckBox()
        cbInegalitaire.Text = "Inégalitaire"
        cbInegalitaire.Location = New Point(450, y4)
        cbInegalitaire.Size = New Size(110, 23)
        grpOeuvre.Controls.Add(cbInegalitaire)

        ' Ligne 5 : Déclaration / Format
        grpOeuvre.Controls.Add(MkLabel("Déclaration :", 8, y5))
        cbDeclaration = New ComboBox()
        cbDeclaration.Location = New Point(90, y5)
        cbDeclaration.Size = New Size(460, 23)
        cbDeclaration.DropDownStyle = ComboBoxStyle.DropDown
        grpOeuvre.Controls.Add(cbDeclaration)
        grpOeuvre.Controls.Add(MkLabel("Format :", 565, y5))
        cbFormat = New ComboBox()
        cbFormat.Location = New Point(620, y5)
        cbFormat.Size = New Size(488, 23)
        cbFormat.DropDownStyle = ComboBoxStyle.DropDown
        grpOeuvre.Controls.Add(cbFormat)

        ' Ligne 6 : Fait à / Fait le
        grpOeuvre.Controls.Add(MkLabel("Fait à :", 8, y6))
        txtFaita = MkTextBox(60, y6, 200)
        grpOeuvre.Controls.Add(txtFaita)
        grpOeuvre.Controls.Add(MkLabel("Fait le :", 275, y6))
        dtFaitle = New DateTimePicker()
        dtFaitle.Location = New Point(325, y6)
        dtFaitle.Size = New Size(110, 23)
        dtFaitle.Format = DateTimePickerFormat.Custom
        dtFaitle.CustomFormat = "dd/MM/yyyy"
        grpOeuvre.Controls.Add(dtFaitle)

        Me.Controls.Add(grpOeuvre)

        ' ── GroupBox AYANTS DROIT ────────────────────────────────
        Dim grpAD As New GroupBox()
        grpAD.Text = "Ayants droit"
        grpAD.Location = New Point(10, 220)
        grpAD.Size = New Size(1120, 410)
        grpAD.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        ' Barre de sélection
        cbPersonnes = New ComboBox()
        cbPersonnes.Location = New Point(10, 25)
        cbPersonnes.Size = New Size(380, 23)
        grpAD.Controls.Add(cbPersonnes)

        btnAjouter = New Button()
        btnAjouter.Text = "AJOUTER"
        btnAjouter.Location = New Point(400, 24)
        btnAjouter.Size = New Size(85, 25)
        grpAD.Controls.Add(btnAjouter)

        btnCalculer = New Button()
        btnCalculer.Text = "CALCULER %"
        btnCalculer.Location = New Point(495, 24)
        btnCalculer.Size = New Size(95, 25)
        grpAD.Controls.Add(btnCalculer)

        btnGererFiches = New Button()
        btnGererFiches.Text = "Gérer les fiches"
        btnGererFiches.Location = New Point(600, 24)
        btnGererFiches.Size = New Size(120, 25)
        btnGererFiches.BackColor = Color.FromArgb(107, 60, 157)
        btnGererFiches.ForeColor = Color.White
        btnGererFiches.FlatStyle = FlatStyle.Flat
        grpAD.Controls.Add(btnGererFiches)


        btnChargerJson = New Button()
        btnChargerJson.Text = "Charger JSON"
        btnChargerJson.Location = New Point(860, 24)
        btnChargerJson.Size = New Size(110, 25)
        grpAD.Controls.Add(btnChargerJson)

        btnSauvegarder = New Button()
        btnSauvegarder.Text = "Sauvegarder JSON"
        btnSauvegarder.Location = New Point(980, 24)
        btnSauvegarder.Size = New Size(140, 25)
        btnSauvegarder.Font = New Font(btnSauvegarder.Font, FontStyle.Bold)
        btnSauvegarder.BackColor = Color.FromArgb(0, 120, 212)
        btnSauvegarder.ForeColor = Color.White
        btnSauvegarder.FlatStyle = FlatStyle.Flat
        grpAD.Controls.Add(btnSauvegarder)

        ' DataGridView
        cmsGrille = New ContextMenuStrip()
        mnuSupprimer = New ToolStripMenuItem("Supprimer la ligne")
        cmsGrille.Items.Add(mnuSupprimer)

        dgv = New DataGridView()
        dgv.Location = New Point(10, 58)
        dgv.Size = New Size(1098, 310)
        dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        dgv.ContextMenuStrip = cmsGrille
        dgv.AllowUserToAddRows = False
        dgv.ReadOnly = False
        dgv.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
        dgv.RowTemplate.Height = 24
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgv.AllowDrop = True
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grpAD.Controls.Add(dgv)

        ' Statut
        lblStatut = New Label()
        lblStatut.Location = New Point(10, 375)
        lblStatut.Size = New Size(1098, 22)
        lblStatut.Text = "Prêt."
        lblStatut.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        grpAD.Controls.Add(lblStatut)

        Me.Controls.Add(grpAD)

        ' Événements
        AddHandler btnAjouter.Click, AddressOf BtnAjouter_Click
        AddHandler btnCalculer.Click, AddressOf BtnCalculer_Click
        AddHandler btnSauvegarder.Click, AddressOf BtnSauvegarder_Click
        AddHandler btnChargerJson.Click, AddressOf BtnChargerJson_Click
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
        AddHandler Me.Load, AddressOf JsonEditorForm_Load
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CHARGEMENT DU FORMULAIRE
    ' ─────────────────────────────────────────────────────────────
    Private Sub JsonEditorForm_Load(sender As Object, e As EventArgs)
        InitDataTable()
        dgv.DataSource = DtDepotCreateur
        ConfigureGridColumns()
        ChargerGoogleSheet()
    End Sub

    Private Sub InitDataTable()
        DtDepotCreateur = New DataTable()
        ' Clé stable vers le XLSX
        DtDepotCreateur.Columns.Add("Id", GetType(String))
        ' Identité
        DtDepotCreateur.Columns.Add("Type", GetType(String))          ' "Physique" / "Moral"
        DtDepotCreateur.Columns.Add("Designation", GetType(String))
        DtDepotCreateur.Columns.Add("Pseudonyme", GetType(String))
        DtDepotCreateur.Columns.Add("Nom", GetType(String))
        DtDepotCreateur.Columns.Add("Prenom", GetType(String))
        DtDepotCreateur.Columns.Add("Genre", GetType(String))
        DtDepotCreateur.Columns.Add("Nele", GetType(String))
        DtDepotCreateur.Columns.Add("Nea", GetType(String))
        DtDepotCreateur.Columns.Add("SocieteGestion", GetType(String))
        ' Personne morale
        DtDepotCreateur.Columns.Add("FormeJuridique", GetType(String))
        DtDepotCreateur.Columns.Add("Capital", GetType(String))
        DtDepotCreateur.Columns.Add("RCS", GetType(String))
        DtDepotCreateur.Columns.Add("Siren", GetType(String))
        DtDepotCreateur.Columns.Add("GenreRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("PrenomRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("NomRepresentant", GetType(String))
        DtDepotCreateur.Columns.Add("FonctionRepresentant", GetType(String))
        ' BDO
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
        ' Adresse
        DtDepotCreateur.Columns.Add("NumVoie", GetType(String))
        DtDepotCreateur.Columns.Add("TypeVoie", GetType(String))
        DtDepotCreateur.Columns.Add("NomVoie", GetType(String))
        DtDepotCreateur.Columns.Add("CP", GetType(String))
        DtDepotCreateur.Columns.Add("Ville", GetType(String))
        DtDepotCreateur.Columns.Add("Pays", GetType(String))
        ' Contact
        DtDepotCreateur.Columns.Add("Mail", GetType(String))
        DtDepotCreateur.Columns.Add("Tel", GetType(String))
        ' Interne (éditeur par défaut depuis sheet)
        DtDepotCreateur.Columns.Add("_EditeurDefaut", GetType(String))
    End Sub

    Private Sub ConfigureGridColumns()
        ' Masquer tout puis afficher les colonnes utiles
        For Each col As DataGridViewColumn In dgv.Columns
            col.Visible = False
        Next

        Dim visible As New Dictionary(Of String, Integer) From {
            {"Id", 70}, {"Designation", 250}, {"Role", 45}, {"Lettrage", 60},
            {"SocieteGestion", 90}, {"Signataire", 70}, {"COAD_IPI", 130},
            {"PH", 60}
        }

        For Each kvp In visible
            If dgv.Columns.Contains(kvp.Key) Then
                dgv.Columns(kvp.Key).Visible = True
                dgv.Columns(kvp.Key).Width = kvp.Value
                dgv.Columns(kvp.Key).HeaderText = kvp.Key
            End If
        Next

        ' En-têtes lisibles
        dgv.Columns("SocieteGestion").HeaderText = "Société"
        dgv.Columns("COAD_IPI").HeaderText = "COAD/IPI"
        dgv.Columns("Signataire").HeaderText = "Signataire"

        ' Remplacer Managelic et Managesub par des ComboBoxColumn
        SetupMoraleComboColumn("Managelic", 150)
        SetupMoraleComboColumn("Managesub", 150)
    End Sub

    Private Sub SetupMoraleComboColumn(colName As String, width As Integer)
        ' Supprimer l'ancienne colonne texte
        If dgv.Columns.Contains(colName) Then
            dgv.Columns.Remove(colName)
        End If

        ' Créer une ComboBoxColumn
        Dim cbCol As New DataGridViewComboBoxColumn()
        cbCol.Name = colName
        cbCol.HeaderText = colName
        cbCol.DataPropertyName = colName
        cbCol.Width = width
        cbCol.Visible = True
        cbCol.FlatStyle = FlatStyle.Flat
        cbCol.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        ' Alimenter avec les désignations des morales + entrée vide
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

    ''' <summary>Recharge les ComboBox Managelic/Managesub après rechargement du sheet.</summary>
    Private Sub RefreshMoraleColumns()
        SetupMoraleComboColumn("Managelic", 150)
        SetupMoraleComboColumn("Managesub", 150)
    End Sub

    ''' <summary>Recharge les ComboBox Déclaration/Format avec les éditeurs (E) de la grille.</summary>
    Public Sub RefreshDeclarationFormat()
        Dim currentDecl As String = cbDeclaration.Text
        Dim currentFmt As String = cbFormat.Text

        cbDeclaration.Items.Clear()
        cbFormat.Items.Clear()
        cbDeclaration.Items.Add("")
        cbFormat.Items.Add("")

        For Each row As DataRow In DtDepotCreateur.Rows
            Dim role As String = row("Role").ToString().Trim().ToUpper()
            If role = "E" Then
                Dim desig As String = row("Designation").ToString().Trim()
                If Not String.IsNullOrEmpty(desig) Then
                    If Not cbDeclaration.Items.Contains(desig) Then cbDeclaration.Items.Add(desig)
                    If Not cbFormat.Items.Contains(desig) Then cbFormat.Items.Add(desig)
                End If
            End If
        Next

        ' Restaurer la valeur courante
        cbDeclaration.Text = currentDecl
        cbFormat.Text = currentFmt
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CHARGEMENT XLSX LOCAL
    ' ─────────────────────────────────────────────────────────────
    Private Sub ChargerGoogleSheet()
        Dim localPath As String = PersonnesForm.DefaultXlsxPath
        Try
            If File.Exists(localPath) Then
                lblStatut.Text = "Chargement de la base de données..."
                Application.DoEvents()
                DtPersonPhy = LoadSheetXlsxLocal(localPath, "PERSONNEPHYSIQUE")
                DtPersonMor = LoadSheetXlsxLocal(localPath, "PERSONNEMORALE")
                PopulateComboBox()
                RefreshMoraleColumns()
                lblStatut.Text = DtPersonPhy.Rows.Count & " physiques, " & DtPersonMor.Rows.Count & " morales chargés."
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
        ' Recharger le Sheet puis synchroniser la grille
        ChargerGoogleSheet()
        SyncGridAvecSheet()
    End Sub

    ''' <summary>
    ''' Synchronise TOUTES les colonnes de la grille depuis le XLSX (source de vérité).
    ''' Physique : recherche par Nom+Prenom ou Pseudonyme.
    ''' Morale   : recherche par Designation (insensible à la casse).
    ''' </summary>
    Private Sub SyncGridAvecSheet()
        ' Mapping XLSX colonne → DtDepotCreateur colonne
        ' Format : {colonne_xlsx, colonne_grille}
        Dim mapPhy As (String, String)() = {
            ("Pseudonyme", "Pseudonyme"),
            ("Nom", "Nom"),
            ("Prenom", "Prenom"),
            ("Genre", "Genre"),
            ("SocieteGestion", "SocieteGestion"),
            ("Num de voie", "NumVoie"),
            ("Type de voie", "TypeVoie"),
            ("Nom de voie", "NomVoie"),
            ("CP", "CP"),
            ("Ville", "Ville"),
            ("Mail", "Mail"),
            ("Tel", "Tel"),
            ("Date de naissance", "DateNaissance"),
            ("Lieu de naissance", "LieuNaissance"),
            ("N Secu", "NSecu")}

        Dim mapMor As (String, String)() = {
            ("Designation", "Designation"),
            ("SocieteGestion", "SocieteGestion"),
            ("Forme Juridique", "FormeJuridique"),
            ("Capital", "Capital"),
            ("RCS", "RCS"),
            ("Siren", "Siren"),
            ("Num de voie", "NumVoie"),
            ("Type de voie", "TypeVoie"),
            ("Nom de voie", "NomVoie"),
            ("CP", "CP"),
            ("Ville", "Ville"),
            ("Mail", "Mail"),
            ("Tel", "Tel"),
            ("Prenom representant", "PrenomRepresentant"),
            ("Nom representant", "NomRepresentant"),
            ("Fonction representant", "FonctionRepresentant")}

        For Each gridRow As DataRow In DtDepotCreateur.Rows
            Dim tp As String = gridRow("Type").ToString()
            Dim sheetRow As DataRow = Nothing

            If tp = "Physique" AndAlso DtPersonPhy IsNot Nothing Then
                Dim nom As String = gridRow("Nom").ToString().Trim().ToUpper()
                Dim prenom As String = gridRow("Prenom").ToString().Trim().ToUpper()
                Dim pseudo As String = gridRow("Pseudonyme").ToString().Trim().ToUpper()
                For Each r As DataRow In DtPersonPhy.Rows
                    Dim rNom As String = SafeStr(r, "Nom").Trim().ToUpper()
                    Dim rPrenom As String = SafeStr(r, "Prenom").Trim().ToUpper()
                    Dim rPseudo As String = SafeStr(r, "Pseudonyme").Trim().ToUpper()
                    If (Not String.IsNullOrEmpty(nom) AndAlso rNom = nom AndAlso rPrenom = prenom) OrElse
                       (Not String.IsNullOrEmpty(pseudo) AndAlso rPseudo = pseudo) Then
                        sheetRow = r : Exit For
                    End If
                Next
                If sheetRow IsNot Nothing Then
                    ' Copier toutes les colonnes mappées
                    For Each mapping As (String, String) In mapPhy
                        Dim xlsCol As String = mapping.Item1
                        Dim gridCol As String = mapping.Item2
                        If DtDepotCreateur.Columns.Contains(gridCol) Then
                            Dim val As String = SafeStr(sheetRow, xlsCol)
                            gridRow(gridCol) = val
                        End If
                    Next
                    ' COAD / IPI — format spécial "COAD : X" ou "IPI : X"
                    Dim coad As String = SafeStr(sheetRow, "COAD")
                    Dim ipi As String = SafeStr(sheetRow, "IPI")
                    If Not String.IsNullOrEmpty(coad) Then
                        gridRow("COAD_IPI") = "COAD : " & coad
                    ElseIf Not String.IsNullOrEmpty(ipi) Then
                        gridRow("COAD_IPI") = "IPI : " & ipi
                    Else
                        gridRow("COAD_IPI") = ""
                    End If
                    ' Reconstruire Designation depuis le XLSX
                    Dim nomXls As String = SafeStr(sheetRow, "Nom").Trim()
                    Dim prenomXls As String = SafeStr(sheetRow, "Prenom").Trim()
                    Dim pseudoXls As String = SafeStr(sheetRow, "Pseudonyme").Trim()
                    If Not String.IsNullOrEmpty(pseudoXls) Then
                        gridRow("Designation") = nomXls & " " & prenomXls & " / " & pseudoXls
                    Else
                        gridRow("Designation") = (nomXls & " " & prenomXls).Trim()
                    End If
                End If

            ElseIf tp = "Moral" AndAlso DtPersonMor IsNot Nothing Then
                Dim desigUp As String = gridRow("Designation").ToString().Trim().ToUpper()
                For Each r As DataRow In DtPersonMor.Rows
                    If SafeStr(r, "Designation").Trim().ToUpper() = desigUp Then
                        sheetRow = r : Exit For
                    End If
                Next
                If sheetRow IsNot Nothing Then
                    ' Copier toutes les colonnes mappées
                    For Each mapping As (String, String) In mapMor
                        Dim xlsCol As String = mapping.Item1
                        Dim gridCol As String = mapping.Item2
                        If DtDepotCreateur.Columns.Contains(gridCol) Then
                            Dim val As String = SafeStr(sheetRow, xlsCol)
                            gridRow(gridCol) = val
                        End If
                    Next
                    ' COAD / IPI — format spécial
                    Dim coad As String = SafeStr(sheetRow, "COAD")
                    Dim ipi As String = SafeStr(sheetRow, "IPI")
                    If Not String.IsNullOrEmpty(coad) Then
                        gridRow("COAD_IPI") = "COAD : " & coad
                    ElseIf Not String.IsNullOrEmpty(ipi) Then
                        gridRow("COAD_IPI") = "IPI : " & ipi
                    Else
                        gridRow("COAD_IPI") = ""
                    End If
                End If
            End If
        Next

        dgv.Refresh()
        lblStatut.Text = "Grille synchronisée avec les fiches."
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
                If Not String.IsNullOrEmpty(desig) Then
                    items.Add($"(Morale) {desig}")
                End If
            Next
        End If

        items.Sort()
        cbPersonnes.Items.AddRange(items.ToArray())
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' AJOUT D'UN AYANT DROIT
    ' ─────────────────────────────────────────────────────────────
    Private Sub BtnAjouter_Click(sender As Object, e As EventArgs)
        Dim sel As String = If(cbPersonnes.SelectedItem IsNot Nothing, cbPersonnes.SelectedItem.ToString(), "")
        If String.IsNullOrEmpty(sel) Then
            MessageBox.Show("Sélectionnez une personne.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Recharger le xlsx pour avoir les données les plus récentes
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

    ''' <summary>Relit le xlsx depuis le disque pour avoir les données fraîches.</summary>
    Private Sub RafraichirBDD()
        Dim localPath As String = PersonnesForm.DefaultXlsxPath
        If Not File.Exists(localPath) Then Return
        Try
            DtPersonPhy = LoadSheetXlsxLocal(localPath, "PERSONNEPHYSIQUE")
            DtPersonMor = LoadSheetXlsxLocal(localPath, "PERSONNEMORALE")
        Catch
            ' Garder les données en mémoire si erreur
        End Try
    End Sub

    Private Sub AjouterPersonneMorale(designation As String)
        If DtPersonMor Is Nothing Then Return
        Dim foundRow As DataRow = DtPersonMor.AsEnumerable().
            FirstOrDefault(Function(r) SafeStr(r, "Designation").Trim().ToUpper() = designation.Trim().ToUpper())
        If foundRow Is Nothing Then Return

        If DtDepotCreateur.Rows.Count = 0 Then
            MessageBox.Show("Ajoutez d'abord un créateur (A ou C) avant un éditeur.", "Information",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Sélection du créateur associé
        Dim createurs As New List(Of String)(
            DtDepotCreateur.AsEnumerable().
            Where(Function(r) r("Role").ToString() = "A" OrElse r("Role").ToString() = "C").
            Select(Function(r) r("Designation").ToString()))

        Using f As New FormCreateurs(createurs)
            If f.ShowDialog() <> DialogResult.OK Then Return
            For Each crea As String In f.SelectedCreateurs
                Dim creaRow As DataRow = DtDepotCreateur.Select($"Designation = '{crea.Replace("'", "''")}'").FirstOrDefault()
                If creaRow Is Nothing Then Continue For

                If String.IsNullOrEmpty(creaRow("Lettrage").ToString()) Then
                    creaRow("Lettrage") = GetNextLetter()
                End If

                Dim nr As DataRow = DtDepotCreateur.NewRow()
                nr("Type") = "Moral"
                nr("Designation") = designation
                nr("Role") = "E"
                nr("Lettrage") = creaRow("Lettrage")
                nr("SocieteGestion") = SafeStr(foundRow, "SocieteGestion", "SACEM")
                nr("Signataire") = True
                CopierAdresseContact(nr, foundRow)
                CopierInfosMorale(nr, foundRow)
                DtDepotCreateur.Rows.Add(nr)
                ' Persistance : mettre à jour la colonne Editeur dans DtPersonPhy
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
            Dim rNom As String = r("Nom").ToString().Trim()
            Dim rPrenom As String = SafeStr(r, "Prenom").Trim()
            Dim rPseudo As String = r("Pseudonyme").ToString().Trim()
            If (rNom = nom AndAlso rPrenom = prenom) OrElse
               (Not String.IsNullOrEmpty(pseudo) AndAlso rPseudo = pseudo) Then
                foundRow = r
                Exit For
            End If
        Next
        If foundRow Is Nothing Then Return

        ' Rôle
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
        nr("Type") = "Physique"
        nr("Id") = SafeStr(foundRow, "Id")
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

        ' Éditeurs par défaut — format "M00001:60;M00002:40" ou "M00001;M00002"
        If Not String.IsNullOrEmpty(editeurDefaut) Then
            ' Parser les entrées et extraire les parts explicites
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
                    If Double.TryParse(tokens(1).Trim(), p) Then
                        partsExplicites(eid) = p
                    End If
                End If
            Next

            ' Si pas de parts explicites → répartition égale
            Dim nbEds As Integer = editeurIds.Count
            For Each eid As String In editeurIds
                If Not partsExplicites.ContainsKey(eid) Then
                    partsExplicites(eid) = Math.Round(100.0 / nbEds, 2)
                End If
            Next

            ' Ajouter chaque éditeur avec sa part calculée
            For Each eid As String In editeurIds
                Dim quotePartCoed As Double = partsExplicites(eid) / 100.0
                AjouterEditeurParDefaut(eid, nr, quotePartCoed)
            Next
        Else
            ' Demander EAC si déjà un éditeur
            If DtDepotCreateur.AsEnumerable().Any(Function(r) r("Role").ToString() = "E") Then
                If MessageBox.Show("Voulez-vous être Éditeur À Compte d'Auteur (EAC) ?",
                                   "EAC ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Dim nrEAC As DataRow = DtDepotCreateur.NewRow()
                    nrEAC("Type") = "Physique"
                    nrEAC("Designation") = nr("Designation").ToString() & " (EAC)"
                    nrEAC("Nom") = nr("Nom")
                    nrEAC("Prenom") = nr("Prenom")
                    nrEAC("Role") = "E"
                    nrEAC("Lettrage") = nr("Lettrage")
                    nrEAC("SocieteGestion") = nr("SocieteGestion")
                    nrEAC("Signataire") = True
                    CopierAdresseContact(nrEAC, foundRow)
                    DtDepotCreateur.Rows.Add(nrEAC)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Ajoute un éditeur lié à un AC dans la grille.
    ''' quotePartCoed = fraction de la part AC attribuée à cet éditeur (ex: 0.6 pour 60%).
    ''' PH de l'éditeur = PH de l'AC * quotePartCoed.
    ''' </summary>
    Private Sub AjouterEditeurParDefaut(editeurId As String, creaRow As DataRow,
                                        Optional quotePartCoed As Double = 1.0)
        ' Calculer PH éditeur = PH AC * quote-part coéditeur
        Dim phAC As Double = 0
        Double.TryParse(creaRow("PH").ToString(), phAC)
        Dim phEditeur As Double = Math.Round(phAC * quotePartCoed, 2)
        Dim phStr As String = If(phEditeur > 0, phEditeur.ToString("0.##"), "")

        If editeurId = "EAC" Then
            Dim nrEAC2 As DataRow = DtDepotCreateur.NewRow()
            nrEAC2("Type") = "Physique"
            nrEAC2("Designation") = creaRow("Designation").ToString() & " (EAC)"
            nrEAC2("Nom") = creaRow("Nom")
            nrEAC2("Prenom") = creaRow("Prenom")
            nrEAC2("Role") = "E"
            nrEAC2("Lettrage") = creaRow("Lettrage")
            nrEAC2("PH") = phStr
            nrEAC2("SocieteGestion") = creaRow("SocieteGestion")
            nrEAC2("Signataire") = True
            DtDepotCreateur.Rows.Add(nrEAC2)
            Return
        End If

        ' Recherche par Id (format M00001) puis fallback Designation
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

    ' ─────────────────────────────────────────────────────────────
    ' CALCUL DES POURCENTAGES
    ' ─────────────────────────────────────────────────────────────
    Private Sub BtnCalculer_Click(sender As Object, e As EventArgs)
        Try
            CalculerPourcentages()
            UpdateLettrages()
            dgv.Refresh()
            lblStatut.Text = "Pourcentages calculés."
        Catch ex As Exception
            MessageBox.Show($"Erreur calcul : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CalculerPourcentages()
        Dim rows = DtDepotCreateur.AsEnumerable().ToList()
        Dim countA As Integer = rows.Where(Function(r) r("Role").ToString() = "A").Count()
        Dim countC As Integer = rows.Where(Function(r) r("Role").ToString() = "C").Count()
        Dim countE As Integer = rows.Where(Function(r) r("Role").ToString() = "E").Count()
        Dim inegal As Boolean = cbInegalitaire.Checked

        For Each row As DataRow In rows
            Dim role As String = row("Role").ToString()
            Dim ph As Decimal = 0
            Decimal.TryParse(row("PH").ToString().Replace(",", "."),
                Globalization.NumberStyles.Any,
                Globalization.CultureInfo.InvariantCulture, ph)

            Dim de As Decimal = 0
            Dim dr As Decimal = 0

            If Not inegal Then
                ' Égalitaire
                If countE > 0 Then
                    If countA > 0 AndAlso countC > 0 Then
                        If role = "A" Then de = Math.Round(100D / 3 / countA, 4) : dr = Math.Round(100D / 3 / countA, 4)
                        If role = "C" Then de = Math.Round(100D / 3 / countC, 4) : dr = Math.Round(100D / 3 / countC, 4)
                        If role = "E" Then de = Math.Round(100D / 3 / countE, 4) : dr = Math.Round(100D / 3 / countE, 4)
                    ElseIf countA > 0 Then
                        If role = "A" Then de = Math.Round(200D / 3 / countA, 4) : dr = Math.Round(50D / countA, 4)
                        If role = "E" Then de = Math.Round(100D / 3 / countE, 4) : dr = Math.Round(50D / countE, 4)
                    ElseIf countC > 0 Then
                        If role = "C" Then de = Math.Round(200D / 3 / countC, 4) : dr = Math.Round(50D / countC, 4)
                        If role = "E" Then de = Math.Round(100D / 3 / countE, 4) : dr = Math.Round(50D / countE, 4)
                    End If
                Else
                    If role = "A" Then de = Math.Round(100D / countA, 4) : dr = de
                    If role = "C" Then de = Math.Round(100D / countC, 4) : dr = de
                End If
            End If
            ' Inégalitaire : laisser l'utilisateur saisir PH manuellement pour l'instant

            row("DE") = de.ToString(Globalization.CultureInfo.InvariantCulture)
            row("DR") = dr.ToString(Globalization.CultureInfo.InvariantCulture)
        Next
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' SAUVEGARDE JSON — FORMAT ALLÉGÉ (Ref + BDO uniquement)
    ' ─────────────────────────────────────────────────────────────
    Private Sub BtnSauvegarder_Click(sender As Object, e As EventArgs)
        SaveToJson()
    End Sub

    Private Sub SaveToJson()
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
                Dim tp As String = row("Type").ToString()

                ad("Id") = row("Id").ToString().Trim()
                ad("Role") = row("Role").ToString()
                ad("Lettrage") = row("Lettrage").ToString()
                ad("PH") = row("PH").ToString()
                ad("Managelic") = row("Managelic").ToString()
                ad("Managesub") = row("Managesub").ToString()
                Dim sig As Boolean = True
                Boolean.TryParse(row("Signataire").ToString(), sig)
                ad("Signataire") = sig

                arr.Add(ad)
            Next

            obj("AyantsDroit") = arr

            File.WriteAllText(sfd.FileName, obj.ToString(Newtonsoft.Json.Formatting.Indented),
                              System.Text.Encoding.UTF8)
            SavedJsonPath = sfd.FileName
            lblStatut.Text = "JSON sauvegardé : " & sfd.FileName
            MessageBox.Show("JSON sauvegardé avec succès !" & vbCrLf & sfd.FileName,
                            "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CHARGEMENT JSON EXISTANT
    ' ─────────────────────────────────────────────────────────────
    Private Sub BtnChargerJson_Click(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Fichiers JSON (*.json)|*.json"
            ofd.Title = "Charger un fichier JSON SACEM"
            If ofd.ShowDialog() = DialogResult.OK Then
                LoadJsonData(ofd.FileName)
                ApplyRowColors()
            End If
        End Using
    End Sub

    Private Sub LoadJsonData(filePath As String)
        Try
            ' Utiliser SACEMJsonReader pour charger et enrichir depuis xlsx
            Dim sacemData As SACEMData = SACEMJsonReader.LoadFromFile(filePath)

            ' Champs de l'oeuvre
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

            ' Ayants droit — BDO conservé depuis JSON, identité/adresse/contact réenrichis depuis XLSX
            RafraichirBDD()
            DtDepotCreateur.Rows.Clear()
            For Each ayant As AyantDroit In sacemData.AyantsDroit
                Dim nr As DataRow = DtDepotCreateur.NewRow()
                nr("Type") = ayant.Identite.Type

                ' BDO depuis JSON (propre au dépôt — ne jamais écraser)
                nr("Id") = ayant.BDO.Id
                nr("Role") = ayant.BDO.Role
                nr("Lettrage") = ayant.BDO.Lettrage
                nr("PH") = ayant.BDO.PH
                nr("Managelic") = ayant.BDO.Managelic
                nr("Managesub") = ayant.BDO.Managesub
                nr("Signataire") = ayant.BDO.Signataire

                ' Identité depuis JSON (fallback si non trouvé dans XLSX)
                Dim desigJson As String = If(Not String.IsNullOrEmpty(ayant.Identite.Designation),
                                             ayant.Identite.Designation,
                                             (ayant.Identite.Prenom & " " & ayant.Identite.Nom).Trim())
                nr("Designation") = desigJson
                nr("Pseudonyme") = ayant.Identite.Pseudonyme
                nr("Nom") = ayant.Identite.Nom
                nr("Prenom") = ayant.Identite.Prenom
                nr("Genre") = ayant.Identite.Genre
                nr("SocieteGestion") = If(String.IsNullOrEmpty(ayant.Identite.SocieteGestion), "SACEM", ayant.Identite.SocieteGestion)
                nr("FormeJuridique") = ayant.Identite.FormeJuridique
                nr("Capital") = ayant.Identite.Capital
                nr("RCS") = ayant.Identite.RCS
                nr("Siren") = ayant.Identite.Siren
                nr("PrenomRepresentant") = ayant.Identite.PrenomRepresentant
                nr("NomRepresentant") = ayant.Identite.NomRepresentant
                nr("GenreRepresentant") = ayant.Identite.GenreRepresentant
                nr("FonctionRepresentant") = ayant.Identite.FonctionRepresentant
                nr("Nele") = ayant.Identite.Nele
                nr("Nea") = ayant.Identite.Nea
                nr("NumVoie") = ayant.Adresse.NumVoie
                nr("TypeVoie") = ayant.Adresse.TypeVoie
                nr("NomVoie") = ayant.Adresse.NomVoie
                nr("CP") = ayant.Adresse.CP
                nr("Ville") = ayant.Adresse.Ville
                nr("Pays") = ayant.Adresse.Pays
                nr("Mail") = ayant.Contact.Mail
                nr("Tel") = ayant.Contact.Tel

                ' Réenrichissement depuis XLSX si disponible
                Dim idJson As String = ayant.BDO.Id
                If ayant.Identite.Type = "Physique" AndAlso DtPersonPhy IsNot Nothing Then
                    ' Recherche par Id (stable) puis fallback par nom (compatibilité anciens JSON)
                    Dim phyRow As DataRow = Nothing
                    If Not String.IsNullOrEmpty(idJson) AndAlso DtPersonPhy.Columns.Contains("Id") Then
                        phyRow = DtPersonPhy.AsEnumerable().FirstOrDefault(
                            Function(r) SafeStr(r, "Id").Trim().ToUpper() = idJson.Trim().ToUpper())
                    End If
                    If phyRow Is Nothing Then
                        ' Fallback nom pour anciens JSON sans Id
                        phyRow = DtPersonPhy.AsEnumerable().FirstOrDefault(
                            Function(r)
                                Dim pseudo As String = SafeStr(r, "Pseudonyme").Trim().ToUpper()
                                Dim nomPrenom As String = (SafeStr(r, "Nom") & " " & SafeStr(r, "Prenom")).Trim().ToUpper()
                                Dim prenomNom As String = (SafeStr(r, "Prenom") & " " & SafeStr(r, "Nom")).Trim().ToUpper()
                                Dim target As String = desigJson.ToUpper()
                                Return pseudo = target OrElse nomPrenom = target OrElse prenomNom = target
                            End Function)
                    End If
                    ' Marquer si Id inconnu (personne supprimée ou non encore dans XLSX)
                    nr("_Orphelin") = (phyRow Is Nothing AndAlso Not String.IsNullOrEmpty(idJson))
                    If phyRow IsNot Nothing Then
                        ' Réenrichir adresse, contact, identité civile
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
                        ' Recalculer Designation avec les données fraîches
                        nr("Designation") = BuildDesignation(phyRow)
                    End If
                ElseIf ayant.Identite.Type = "Moral" AndAlso DtPersonMor IsNot Nothing Then
                    Dim morRow As DataRow = Nothing
                    If Not String.IsNullOrEmpty(idJson) AndAlso DtPersonMor.Columns.Contains("Id") Then
                        morRow = DtPersonMor.AsEnumerable().FirstOrDefault(
                            Function(r) SafeStr(r, "Id").Trim().ToUpper() = idJson.Trim().ToUpper())
                    End If
                    If morRow Is Nothing Then
                        morRow = DtPersonMor.AsEnumerable().FirstOrDefault(
                            Function(r) SafeStr(r, "Designation").Trim().ToUpper() = desigJson.ToUpper())
                    End If
                    nr("_Orphelin") = (morRow Is Nothing AndAlso Not String.IsNullOrEmpty(idJson))
                    If morRow IsNot Nothing Then
                        ' Réenrichir infos légales et contact
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
                    End If
                End If

                DtDepotCreateur.Rows.Add(nr)
            Next

            dgv.Refresh()
            lblStatut.Text = "JSON chargé : " & Path.GetFileName(filePath) & " — données enrichies depuis BDD"
            RefreshDeclarationFormat()
        Catch ex As Exception
            MessageBox.Show("Erreur chargement JSON : " & ex.Message, "Erreur",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' GRILLE — ÉVÉNEMENTS
    ' ─────────────────────────────────────────────────────────────
    Private Sub Dgv_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs)
        Dim colName As String = dgv.Columns(e.ColumnIndex).Name
        If colName = "Role" Then
            Dim val As String = e.FormattedValue.ToString().Trim().ToUpper()
            Dim valid As String() = {"A", "C", "AR", "AD", "E"}
            If Not valid.Contains(val) Then
                e.Cancel = True
                MessageBox.Show("Rôles valides : A, C, AR, AD, E", "Valeur invalide",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub Dgv_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        Dim editable As String() = {"Role", "Lettrage", "PH", "SocieteGestion", "Signataire", "Managelic", "Managesub", "COAD_IPI"}
        If Not editable.Contains(dgv.Columns(e.ColumnIndex).Name) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub Dgv_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        ' Ignorer les erreurs de valeur invalide dans les ComboBox (valeur absente de la liste)
        ' La valeur sera quand même conservée dans le DataTable
        e.ThrowException = False
        ' Si c'est Managelic ou Managesub, ajouter la valeur manquante à la liste
        Dim colName As String = dgv.Columns(e.ColumnIndex).Name
        If colName = "Managelic" OrElse colName = "Managesub" Then
            Dim cbCol As DataGridViewComboBoxColumn = TryCast(dgv.Columns(e.ColumnIndex), DataGridViewComboBoxColumn)
            If cbCol IsNot Nothing Then
                Dim rowVal As String = ""
                If e.RowIndex >= 0 AndAlso e.RowIndex < DtDepotCreateur.Rows.Count Then
                    rowVal = DtDepotCreateur.Rows(e.RowIndex)(colName).ToString()
                End If
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
                If r("Lettrage").ToString() = lettrage AndAlso r("Role").ToString() = "E" Then
                    DtDepotCreateur.Rows.Remove(r)
                End If
            Next
        End If

        dgv.Rows.Remove(selRow)
        UpdateLettrages()
        RefreshDeclarationFormat()
        dgv.Refresh()
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' DRAG & DROP DES LIGNES
    ' ─────────────────────────────────────────────────────────────
    Private Sub Dgv_MouseDown(sender As Object, e As MouseEventArgs)
        Dim hit = dgv.HitTest(e.X, e.Y)
        If hit.RowIndex >= 0 Then
            _dragRowIndex = hit.RowIndex
        End If
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

        ' Copier la ligne source
        Dim sourceRow As DataRow = DtDepotCreateur.Rows(sourceIndex)
        Dim newRow As DataRow = DtDepotCreateur.NewRow()
        newRow.ItemArray = sourceRow.ItemArray.Clone()

        ' Supprimer la source et réinsérer à destination
        DtDepotCreateur.Rows.RemoveAt(sourceIndex)
        If destIndex > DtDepotCreateur.Rows.Count Then
            DtDepotCreateur.Rows.Add(newRow)
        Else
            DtDepotCreateur.Rows.InsertAt(newRow, destIndex)
        End If

        dgv.ClearSelection()
        If destIndex < dgv.Rows.Count Then
            dgv.Rows(destIndex).Selected = True
        End If

        ApplyRowColors()
        RefreshDeclarationFormat()
        _dragRowIndex = -1
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
                ' Id du JSON introuvable dans le XLSX (personne supprimée ou BDD absente)
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
        ' Cherche d'abord le nom exact, puis la variante normalisée
        If row.Table.Columns.Contains(colName) AndAlso Not IsDBNull(row(colName)) Then
            Return row(colName).ToString()
        End If
        ' Fallback : chercher colonne avec nom normalisé (sans accents)
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
        ' COAD / IPI
        Dim coad As String = SafeStr(source, "COAD")
        Dim ipi As String = SafeStr(source, "IPI")
        If Not String.IsNullOrEmpty(coad) Then
            dest("COAD_IPI") = "COAD : " & coad
        ElseIf Not String.IsNullOrEmpty(ipi) Then
            dest("COAD_IPI") = "IPI : " & ipi
        End If
    End Sub

    Private Function AjusterRole(role As String, genre As String) As String
        Dim litteraires As String() = {"Billet d'humeur", "Texte", "Texte de présentation",
                                       "Texte de sketch", "Poeme", "Chronique"}
        Dim musicaux As String() = {"Instrumental", "Musique illustrative"}
        If litteraires.Contains(genre) AndAlso role = "C" Then Return "A"
        If musicaux.Contains(genre) AndAlso role = "A" Then Return "C"
        Return role
    End Function

    Private Function RolesPossibles(genre As String) As List(Of String)
        Dim litteraires As String() = {"Billet d'humeur", "Texte", "Texte de présentation",
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

    Private Function MkLabel(text As String, x As Integer, y As Integer) As Label
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Location = New Point(x, y + 3)
        lbl.AutoSize = True
        Return lbl
    End Function

    Private Function MkTextBox(x As Integer, y As Integer, width As Integer) As TextBox
        Dim tb As New TextBox()
        tb.Location = New Point(x, y)
        tb.Size = New Size(width, 23)
        Return tb
    End Function


    Private Sub Dgv_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim row As DataRow = DtDepotCreateur.Rows(e.RowIndex)
        Dim tp As String = row("Type").ToString().Trim()

        If tp = "Physique" Then
            OuvrirFichePhysique(row)
        ElseIf tp = "Moral" Then
            OuvrirFicheMorale(row)
        End If
    End Sub

    Private Sub OuvrirFichePhysique(row As DataRow)
        ' Retrouver la ligne complete dans DtPersonPhy (source = XLSX)
        Dim srcRow As DataRow = Nothing
        If DtPersonPhy IsNot Nothing Then
            Dim nom As String = row("Nom").ToString().Trim()
            Dim prenom As String = row("Prenom").ToString().Trim()
            Dim pseudo As String = row("Pseudonyme").ToString().Trim()
            For Each r As DataRow In DtPersonPhy.Rows
                Dim rNom As String = SafeStr(r, "Nom").Trim()
                Dim rPrenom As String = SafeStr(r, "Prenom").Trim()
                Dim rPseudo As String = SafeStr(r, "Pseudonyme").Trim()
                If (Not String.IsNullOrEmpty(nom) AndAlso rNom = nom AndAlso rPrenom = prenom) OrElse
                   (Not String.IsNullOrEmpty(pseudo) AndAlso rPseudo = pseudo) Then
                    srcRow = r
                    Exit For
                End If
            Next
        End If

        ' Construire le tableau de valeurs depuis srcRow (XLSX) ou row (grille) en fallback
        Dim cols() As String = PersonnesForm.ColsPhy
        Dim vals(cols.Length - 1) As String
        For i As Integer = 0 To cols.Length - 1
            Dim col As String = cols(i)
            If srcRow IsNot Nothing Then
                ' Chercher la colonne dans DtPersonPhy (nom exact ou variante)
                Try
                    vals(i) = srcRow(col).ToString()
                Catch
                    vals(i) = ""
                End Try
            End If
            ' Fallback sur DtDepotCreateur pour les champs qu on a
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

            ' Mettre à jour DtPersonPhy si la ligne source existe
            If srcRow IsNot Nothing Then
                For i As Integer = 0 To cols.Length - 1
                    Try
                        srcRow(cols(i)) = res(i)
                    Catch
                    End Try
                Next
            End If

            ' Mettre à jour DtDepotCreateur (grille) — résolution par nom de colonne
            Dim idxPseudo As Integer = Array.IndexOf(cols, "Pseudonyme")
            Dim idxNom As Integer = Array.IndexOf(cols, "Nom")
            Dim idxPrenom As Integer = Array.IndexOf(cols, "Prenom")
            Dim idxGenre As Integer = Array.IndexOf(cols, "Genre")
            Dim idxRole As Integer = Array.IndexOf(cols, "Role")
            Dim idxCOAD As Integer = Array.IndexOf(cols, "COAD")
            Dim idxIPI As Integer = Array.IndexOf(cols, "IPI")
            Dim idxNumV As Integer = Array.IndexOf(cols, "Num de voie")
            Dim idxTypeV As Integer = Array.IndexOf(cols, "Type de voie")
            Dim idxNomV As Integer = Array.IndexOf(cols, "Nom de voie")
            Dim idxCP As Integer = Array.IndexOf(cols, "CP")
            Dim idxVille As Integer = Array.IndexOf(cols, "Ville")
            Dim idxMail As Integer = Array.IndexOf(cols, "Mail")
            Dim idxTel As Integer = Array.IndexOf(cols, "Tel")
            Dim idxId As Integer = Array.IndexOf(cols, "Id")

            Dim pseudo2 As String = If(idxPseudo >= 0, res(idxPseudo).Trim(), "")
            Dim nom2 As String = If(idxNom >= 0, res(idxNom).Trim(), "")
            Dim prenom2 As String = If(idxPrenom >= 0, res(idxPrenom).Trim(), "")

            If idxPseudo >= 0 Then row("Pseudonyme") = pseudo2
            If idxNom >= 0 Then row("Nom") = nom2
            If idxPrenom >= 0 Then row("Prenom") = prenom2
            If idxGenre >= 0 Then row("Genre") = res(idxGenre)
            If idxRole >= 0 Then row("Role") = res(idxRole)
            If srcRow IsNot Nothing Then row("Id") = SafeStr(srcRow, "Id")

            Dim coad As String = If(idxCOAD >= 0, res(idxCOAD).Trim(), "")
            Dim ipi As String = If(idxIPI >= 0, res(idxIPI).Trim(), "")
            If Not String.IsNullOrEmpty(coad) Then
                row("COAD_IPI") = "COAD : " & coad
            ElseIf Not String.IsNullOrEmpty(ipi) Then
                row("COAD_IPI") = "IPI : " & ipi
            End If

            If idxNumV >= 0 Then row("NumVoie") = res(idxNumV)
            If idxTypeV >= 0 Then row("TypeVoie") = res(idxTypeV)
            If idxNomV >= 0 Then row("NomVoie") = res(idxNomV)
            If idxCP >= 0 Then row("CP") = res(idxCP)
            If idxVille >= 0 Then row("Ville") = res(idxVille)
            If idxMail >= 0 Then row("Mail") = res(idxMail)
            If idxTel >= 0 Then row("Tel") = res(idxTel)

            ' Reconstruire Designation
            If Not String.IsNullOrEmpty(pseudo2) Then
                row("Designation") = nom2 & " " & prenom2 & " / " & pseudo2
            Else
                row("Designation") = (nom2 & " " & prenom2).Trim()
            End If
            dgv.Refresh()
            ApplyRowColors()
            lblStatut.Text = "Personne physique mise à jour."
        End Using
    End Sub

    Private Sub OuvrirFicheMorale(row As DataRow)
        ' Retrouver la ligne complete dans DtPersonMor (source = XLSX)
        Dim srcRow As DataRow = Nothing
        If DtPersonMor IsNot Nothing Then
            Dim desig As String = row("Designation").ToString().Trim().ToUpper()
            For Each r As DataRow In DtPersonMor.Rows
                If SafeStr(r, "Designation").Trim().ToUpper() = desig Then
                    srcRow = r
                    Exit For
                End If
            Next
        End If

        Dim cols() As String = PersonnesForm.ColsMor
        Dim vals(cols.Length - 1) As String
        For i As Integer = 0 To cols.Length - 1
            Dim col As String = cols(i)
            If srcRow IsNot Nothing Then
                Try
                    vals(i) = srcRow(col).ToString()
                Catch
                    vals(i) = ""
                End Try
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

            ' Mettre à jour DtPersonMor si la ligne source existe
            If srcRow IsNot Nothing Then
                For i As Integer = 0 To cols.Length - 1
                    Try
                        srcRow(cols(i)) = res(i)
                    Catch
                    End Try
                Next
            End If

            ' Mettre à jour DtDepotCreateur (grille) — résolution par nom de colonne
            Dim idxDesig As Integer = Array.IndexOf(cols, "Designation")
            Dim idxIdM As Integer = Array.IndexOf(cols, "Id")
            Dim idxCOADm As Integer = Array.IndexOf(cols, "COAD")
            Dim idxIPIm As Integer = Array.IndexOf(cols, "IPI")
            Dim idxFJ As Integer = Array.IndexOf(cols, "Forme Juridique")
            Dim idxCap As Integer = Array.IndexOf(cols, "Capital")
            Dim idxRCS As Integer = Array.IndexOf(cols, "RCS")
            Dim idxSiren As Integer = Array.IndexOf(cols, "Siren")
            Dim idxNumVm As Integer = Array.IndexOf(cols, "Num de voie")
            Dim idxTypeVm As Integer = Array.IndexOf(cols, "Type de voie")
            Dim idxNomVm As Integer = Array.IndexOf(cols, "Nom de voie")
            Dim idxCPm As Integer = Array.IndexOf(cols, "CP")
            Dim idxVillem As Integer = Array.IndexOf(cols, "Ville")
            Dim idxPrenR As Integer = Array.IndexOf(cols, "Prenom representant")
            Dim idxNomR As Integer = Array.IndexOf(cols, "Nom representant")
            Dim idxFoncR As Integer = Array.IndexOf(cols, "Fonction representant")
            Dim idxMailm As Integer = Array.IndexOf(cols, "Mail")
            Dim idxTelm As Integer = Array.IndexOf(cols, "Tel")

            If idxDesig >= 0 Then row("Designation") = res(idxDesig)
            If srcRow IsNot Nothing Then row("Id") = SafeStr(srcRow, "Id")

            Dim coad As String = If(idxCOADm >= 0, res(idxCOADm).Trim(), "")
            Dim ipi As String = If(idxIPIm >= 0, res(idxIPIm).Trim(), "")
            If Not String.IsNullOrEmpty(coad) Then
                row("COAD_IPI") = "COAD : " & coad
            ElseIf Not String.IsNullOrEmpty(ipi) Then
                row("COAD_IPI") = "IPI : " & ipi
            End If

            If idxFJ >= 0 Then row("FormeJuridique") = res(idxFJ)
            If idxCap >= 0 Then row("Capital") = res(idxCap)
            If idxRCS >= 0 Then row("RCS") = res(idxRCS)
            If idxSiren >= 0 Then row("Siren") = res(idxSiren)
            If idxNumVm >= 0 Then row("NumVoie") = res(idxNumVm)
            If idxTypeVm >= 0 Then row("TypeVoie") = res(idxTypeVm)
            If idxNomVm >= 0 Then row("NomVoie") = res(idxNomVm)
            If idxCPm >= 0 Then row("CP") = res(idxCPm)
            If idxVillem >= 0 Then row("Ville") = res(idxVillem)
            If idxPrenR >= 0 Then row("PrenomRepresentant") = res(idxPrenR)
            If idxNomR >= 0 Then row("NomRepresentant") = res(idxNomR)
            If idxFoncR >= 0 Then row("FonctionRepresentant") = res(idxFoncR)
            If idxMailm >= 0 Then row("Mail") = res(idxMailm)
            If idxTelm >= 0 Then row("Tel") = res(idxTelm)
            dgv.Refresh()
            ApplyRowColors()
            lblStatut.Text = "Personne morale mise à jour."
        End Using
    End Sub



End Class

' ─────────────────────────────────────────────────────────────────────────────
' FORMULAIRE SÉLECTION CRÉATEURS (remplace FormCreators de visualseize)
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
        Me.Text = "Sélectionner les créateurs associés"
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
' FORMULAIRE SÉLECTION RÔLE (remplace FormRoles de visualseize)
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
        Me.Text = "Sélectionner le rôle"
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
            MessageBox.Show("Selectionnez un role.", "Attention",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

End Class
