Imports System.IO
Imports System.Windows.Forms
Imports OfficeOpenXml

''' <summary>
''' Formulaire de gestion des fiches Personnes Physiques et Morales.
''' Lecture/écriture dans un fichier xlsx local.
''' </summary>
Public Class PersonnesForm
    Inherits Form

    ' ─────────────────────────────────────────────────────────────
    ' CONSTANTES
    ' ─────────────────────────────────────────────────────────────
    Public Shared ReadOnly DefaultXlsxPath As String =
        Path.Combine(Application.StartupPath, "..", "..", "Data", "BDD_IQS_Person.xlsx")

    ''' <summary>
    ''' À appeler au démarrage de l'application.
    ''' Crée le dossier Data\ si besoin.
    ''' </summary>
    Public Shared Sub InitialiserBDD()
        Try
            Dim dir As String = Path.GetDirectoryName(DefaultXlsxPath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
        Catch
        End Try
    End Sub


    Private Const SHEET_PHY As String = "PERSONNEPHYSIQUE"
    Private Const SHEET_MOR As String = "PERSONNEMORALE"

    ' Colonnes PERSONNEPHYSIQUE
    Public Shared ReadOnly ColsPhy As String() = {
        "Id", "Pseudonyme", "Nom", "Prénom", "Genre", "SocieteGestion",
        "Editeur", "Rôle", "COAD", "IPI", "IPI 2",
        "Num de voie", "Type de voie", "Nom de voie", "CP", "Ville", "Pays",
        "Mail", "Tél", "Date de naissance", "Lieu de naissance", "N° Sécu"
    }

    ' Colonnes PERSONNEMORALE
    Public Shared ReadOnly ColsMor As String() = {
        "Id", "Designation", "SocieteGestion", "COAD", "IPI",
        "Forme Juridique", "Capital", "RCS", "Siren",
        "Num de voie", "Type de voie", "Nom de voie", "CP", "Ville", "Pays",
        "Prénom representant", "Nom representant", "Fonction representant",
        "Mail", "Tél"
    }

    ' ─────────────────────────────────────────────────────────────
    ' DONNÉES
    ' ─────────────────────────────────────────────────────────────
    Private _xlsxPath As String
    Public DtPhy As DataTable
    Public DtMor As DataTable
    Private _modifie As Boolean = False

    ' ─────────────────────────────────────────────────────────────
    ' CONTRÔLES
    ' ─────────────────────────────────────────────────────────────
    Private tabControl As TabControl
    Private tabPhy As TabPage
    Private tabMor As TabPage
    Private dgvPhy As DataGridView
    Private dgvMor As DataGridView
    Private txtSearchPhy As TextBox
    Private txtSearchMor As TextBox
    Private btnAddPhy As Button
    Private btnEditPhy As Button
    Private btnDelPhy As Button
    Private btnAddMor As Button
    Private btnEditMor As Button
    Private btnDelMor As Button
    Private btnSave As Button
    Private lblInfo As Label

    ' ─────────────────────────────────────────────────────────────
    ' CONSTRUCTEUR
    ' ─────────────────────────────────────────────────────────────
    Public Sub New(xlsxPath As String)
        _xlsxPath = xlsxPath
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Gestion des Ayants Droit"
        Me.Size = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimumSize = New Size(900, 500)

        ' TabControl
        tabControl = New TabControl()
        tabControl.Location = New Point(10, 10)
        tabControl.Size = New Size(1165, 600)
        tabControl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        Me.Controls.Add(tabControl)

        ' Onglet Physique
        tabPhy = New TabPage("Personnes Physiques")
        tabControl.TabPages.Add(tabPhy)
        BuildTab(tabPhy, True)

        ' Onglet Morale
        tabMor = New TabPage("Personnes Morales")
        tabControl.TabPages.Add(tabMor)
        BuildTab(tabMor, False)

        ' Barre du bas
        btnSave = New Button()
        btnSave.Text = "Sauvegarder"
        btnSave.Location = New Point(10, 620)
        btnSave.Size = New Size(130, 30)
        btnSave.BackColor = Color.FromArgb(0, 120, 212)
        btnSave.ForeColor = Color.White
        btnSave.FlatStyle = FlatStyle.Flat
        btnSave.Font = New Font(btnSave.Font, FontStyle.Bold)
        btnSave.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Me.Controls.Add(btnSave)


        lblInfo = New Label()
        lblInfo.Location = New Point(160, 626)
        lblInfo.Size = New Size(850, 20)
        lblInfo.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(lblInfo)

        AddHandler Me.Load, AddressOf PersonnesForm_Load
        AddHandler Me.FormClosing, AddressOf PersonnesForm_Closing
        AddHandler btnSave.Click, AddressOf BtnSave_Click
    End Sub

    Private Sub BuildTab(tab As TabPage, isPhy As Boolean)
        ' Barre du haut
        Dim pnlTop As New Panel()
        pnlTop.Dock = DockStyle.Top
        pnlTop.Height = 35
        tab.Controls.Add(pnlTop)

        Dim lblSearch As New Label()
        lblSearch.Text = "Rechercher :"
        lblSearch.Location = New Point(5, 9)
        lblSearch.AutoSize = True
        pnlTop.Controls.Add(lblSearch)

        Dim txtSearch As New TextBox()
        txtSearch.Location = New Point(85, 5)
        txtSearch.Size = New Size(300, 23)
        pnlTop.Controls.Add(txtSearch)

        Dim bAdd As Button = MkBtn("+ Ajouter", 400, 4, 90, Color.FromArgb(16, 124, 16))
        Dim bEdit As Button = MkBtn("Modifier", 500, 4, 80, Color.FromArgb(0, 99, 177))
        Dim bDel As Button = MkBtn("Supprimer", 590, 4, 85, Color.FromArgb(196, 43, 28))
        pnlTop.Controls.Add(bAdd)
        pnlTop.Controls.Add(bEdit)
        pnlTop.Controls.Add(bDel)

        Dim grid As New DataGridView()
        grid.Dock = DockStyle.Fill
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.ReadOnly = True
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.MultiSelect = False
        grid.RowTemplate.Height = 24
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        grid.ScrollBars = ScrollBars.Both
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)
        tab.Controls.Add(grid)

        If isPhy Then
            dgvPhy = grid
            txtSearchPhy = txtSearch
            btnAddPhy = bAdd
            btnEditPhy = bEdit
            btnDelPhy = bDel
            AddHandler txtSearch.TextChanged, AddressOf TxtSearchPhy_Changed
            AddHandler bAdd.Click, AddressOf BtnAddPhy_Click
            AddHandler bEdit.Click, AddressOf BtnEditPhy_Click
            AddHandler bDel.Click, AddressOf BtnDelPhy_Click
            AddHandler grid.CellDoubleClick, AddressOf DgvPhy_DoubleClick
        Else
            dgvMor = grid
            txtSearchMor = txtSearch
            btnAddMor = bAdd
            btnEditMor = bEdit
            btnDelMor = bDel
            AddHandler txtSearch.TextChanged, AddressOf TxtSearchMor_Changed
            AddHandler bAdd.Click, AddressOf BtnAddMor_Click
            AddHandler bEdit.Click, AddressOf BtnEditMor_Click
            AddHandler bDel.Click, AddressOf BtnDelMor_Click
            AddHandler grid.CellDoubleClick, AddressOf DgvMor_DoubleClick
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CHARGEMENT
    ' ─────────────────────────────────────────────────────────────
    Private Sub PersonnesForm_Load(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            If File.Exists(_xlsxPath) Then
                Using pkg As New ExcelPackage(New FileInfo(_xlsxPath))
                    DtPhy = ReadSheet(pkg, SHEET_PHY, ColsPhy)
                    DtMor = ReadSheet(pkg, SHEET_MOR, ColsMor)
                End Using
            Else
                DtPhy = CreateDt(ColsPhy)
                DtMor = CreateDt(ColsMor)
            End If
            dgvPhy.DataSource = DtPhy
            dgvMor.DataSource = DtMor
            SetColWidths(dgvPhy)
            SetColWidths(dgvMor)
            lblInfo.Text = DtPhy.Rows.Count & " personnes physiques, " & DtMor.Rows.Count & " morales."
        Catch ex As Exception
            MessageBox.Show("Erreur chargement : " & ex.Message, "Erreur",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ReadSheet(pkg As ExcelPackage, sheetName As String, cols As String()) As DataTable
        Dim dt As DataTable = CreateDt(cols)
        Dim ws = pkg.Workbook.Worksheets(sheetName)
        If ws Is Nothing OrElse ws.Dimension Is Nothing Then Return dt

        Dim colMap As New Dictionary(Of Integer, Integer)()
        For c = 1 To ws.Dimension.Columns
            Dim header As String = ws.Cells(1, c).Text.Trim()
            ' Normaliser pour comparaison souple
            Dim headerNorm As String = NormalizeHeader(header)
            For i = 0 To cols.Length - 1
                If String.Compare(headerNorm, NormalizeHeader(cols(i)), StringComparison.OrdinalIgnoreCase) = 0 Then
                    colMap(c) = i
                    Exit For
                End If
            Next
        Next

        For r = 2 To ws.Dimension.Rows
            Dim nr As DataRow = dt.NewRow()
            Dim hasData As Boolean = False
            For Each kvp In colMap
                Dim val As String = ws.Cells(r, kvp.Key).Text.Trim()
                nr(kvp.Value) = val
                If Not String.IsNullOrEmpty(val) Then hasData = True
            Next
            If hasData Then dt.Rows.Add(nr)
        Next
        Return dt
    End Function

    Private Function NormalizeHeader(s As String) As String
        ' Supprime accents et espaces pour comparaison souple
        s = s.Trim().ToLower()
        s = s.Replace("é", "e").Replace("è", "e").Replace("ê", "e")
        s = s.Replace("à", "a").Replace("â", "a")
        s = s.Replace("ô", "o").Replace("î", "i").Replace("û", "u")
        s = s.Replace("ç", "c")
        s = s.Replace("°", "").Replace("n°", "n")
        Return s
    End Function

    Private Function CreateDt(cols As String()) As DataTable
        Dim dt As New DataTable()
        For Each col As String In cols
            dt.Columns.Add(col, GetType(String))
        Next
        Return dt
    End Function

    Private Sub SetColWidths(grid As DataGridView)
        Dim widths As New Dictionary(Of String, Integer) From {
            {"Id", 72}, {"Pseudonyme", 110}, {"Nom", 110}, {"Prénom", 95}, {"Prenom", 95},
            {"Genre", 50}, {"SocieteGestion", 80},
            {"Editeur", 180}, {"Rôle", 45}, {"Role", 45},
            {"COAD", 80}, {"IPI", 100}, {"IPI 2", 80},
            {"Designation", 210}, {"Forme Juridique", 100}, {"Capital", 70},
            {"RCS", 80}, {"Siren", 95},
            {"Num de voie", 60}, {"Type de voie", 80}, {"Nom de voie", 150},
            {"CP", 60}, {"Ville", 110}, {"Pays", 80},
            {"Prénom representant", 120}, {"Prenom representant", 120},
            {"Nom representant", 120}, {"Fonction representant", 130},
            {"Mail", 180}, {"Tél", 100}, {"Tel", 100},
            {"Date de naissance", 100}, {"Lieu de naissance", 100},
            {"N° Sécu", 110}, {"N Secu", 110}
        }
        For Each col As DataGridViewColumn In grid.Columns
            If widths.ContainsKey(col.HeaderText) Then
                col.Width = widths(col.HeaderText)
            Else
                col.Width = 100
            End If
        Next
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' RECHERCHE
    ' ─────────────────────────────────────────────────────────────
    Private Sub TxtSearchPhy_Changed(sender As Object, e As EventArgs)
        FiltrerGrille(dgvPhy, DtPhy, txtSearchPhy.Text)
    End Sub

    Private Sub TxtSearchMor_Changed(sender As Object, e As EventArgs)
        FiltrerGrille(dgvMor, DtMor, txtSearchMor.Text)
    End Sub

    Private Sub FiltrerGrille(grid As DataGridView, dt As DataTable, search As String)
        If String.IsNullOrEmpty(search.Trim()) Then
            grid.DataSource = dt
            SetColWidths(grid)
            Return
        End If
        Dim term As String = search.Trim().ToLower()
        Dim dtF As DataTable = dt.Clone()
        For Each row As DataRow In dt.Rows
            For Each cellVal In row.ItemArray
                If cellVal.ToString().ToLower().Contains(term) Then
                    dtF.ImportRow(row)
                    Exit For
                End If
            Next
        Next
        grid.DataSource = dtF
        SetColWidths(grid)
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CRUD PHYSIQUE
    ' ─────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Calcule le prochain ID libre de la forme P00001 ou M00001
    ''' en parcourant tous les IDs existants dans le DataTable.
    ''' </summary>
    Private Function ProchainId(dt As DataTable, prefix As String) As String
        Dim maxNum As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim id As String = If(row(0) IsNot Nothing, row(0).ToString().Trim(), "")
            If id.Length > 1 AndAlso id(0).ToString().ToUpper() = prefix.ToUpper() Then
                Dim numPart As String = id.Substring(1)
                Dim n As Integer
                If Integer.TryParse(numPart, n) AndAlso n > maxNum Then
                    maxNum = n
                End If
            End If
        Next
        Return prefix.ToUpper() & (maxNum + 1).ToString("D5")
    End Function

    Private Sub BtnAddPhy_Click(sender As Object, e As EventArgs)
        Dim initVals(ColsPhy.Length - 1) As String
        initVals(0) = ProchainId(DtPhy, "P")
        Using f As New FichePersonneForm(ColsPhy, initVals, "Nouvelle Personne Physique")
            If f.ShowDialog() = DialogResult.OK Then
                DtPhy.Rows.Add(f.GetValues())
                dgvPhy.DataSource = DtPhy
                SetColWidths(dgvPhy)
                _modifie = True
                lblInfo.Text = "Ligne ajoutée. " & DtPhy.Rows.Count & " personnes physiques."
            End If
        End Using
    End Sub

    Private Sub BtnEditPhy_Click(sender As Object, e As EventArgs)
        EditRow(dgvPhy, DtPhy, ColsPhy, "Modifier Personne Physique")
    End Sub

    Private Sub DgvPhy_DoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        EditRow(dgvPhy, DtPhy, ColsPhy, "Modifier Personne Physique")
    End Sub

    Private Sub BtnDelPhy_Click(sender As Object, e As EventArgs)
        DeleteRow(dgvPhy, DtPhy)
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CRUD MORALE
    ' ─────────────────────────────────────────────────────────────
    Private Sub BtnAddMor_Click(sender As Object, e As EventArgs)
        Dim initVals(ColsMor.Length - 1) As String
        initVals(0) = ProchainId(DtMor, "M")
        ' SocieteGestion (index 2) = SACEM par défaut
        Dim sgIdx As Integer = Array.IndexOf(ColsMor, "SocieteGestion")
        If sgIdx >= 0 Then initVals(sgIdx) = "SACEM"
        Using f As New FichePersonneForm(ColsMor, initVals, "Nouvelle Personne Morale")
            If f.ShowDialog() = DialogResult.OK Then
                DtMor.Rows.Add(f.GetValues())
                dgvMor.DataSource = DtMor
                SetColWidths(dgvMor)
                _modifie = True
                lblInfo.Text = "Ligne ajoutée. " & DtMor.Rows.Count & " personnes morales."
            End If
        End Using
    End Sub

    Private Sub BtnEditMor_Click(sender As Object, e As EventArgs)
        EditRow(dgvMor, DtMor, ColsMor, "Modifier Personne Morale")
    End Sub

    Private Sub DgvMor_DoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        EditRow(dgvMor, DtMor, ColsMor, "Modifier Personne Morale")
    End Sub

    Private Sub BtnDelMor_Click(sender As Object, e As EventArgs)
        DeleteRow(dgvMor, DtMor)
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' HELPERS CRUD
    ' ─────────────────────────────────────────────────────────────
    Private Sub EditRow(grid As DataGridView, dt As DataTable, cols As String(), title As String)
        If grid.SelectedRows.Count = 0 Then
            MessageBox.Show("Selectionnez une ligne.", "Attention",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim selRow As DataGridViewRow = grid.SelectedRows(0)
        Dim firstVal As String = If(selRow.Cells(0).Value IsNot Nothing, selRow.Cells(0).Value.ToString(), "")

        ' Retrouver dans le DataTable source
        Dim dtRow As DataRow = Nothing
        For Each r As DataRow In dt.Rows
            If r(0).ToString() = firstVal Then
                dtRow = r
                Exit For
            End If
        Next
        If dtRow Is Nothing Then Return

        Dim current(cols.Length - 1) As String
        For i = 0 To cols.Length - 1
            current(i) = dtRow(i).ToString()
        Next

        Using f As New FichePersonneForm(cols, current, title)
            If f.ShowDialog() = DialogResult.OK Then
                Dim vals As String() = f.GetValues()
                For i = 0 To cols.Length - 1
                    dtRow(i) = vals(i)
                Next
                grid.DataSource = dt
                SetColWidths(grid)
                _modifie = True
                lblInfo.Text = "Fiche modifiee."
            End If
        End Using
    End Sub

    Private Sub DeleteRow(grid As DataGridView, dt As DataTable)
        If grid.SelectedRows.Count = 0 Then Return
        Dim firstVal As String = If(grid.SelectedRows(0).Cells(0).Value IsNot Nothing,
                                    grid.SelectedRows(0).Cells(0).Value.ToString(), "")

        If MessageBox.Show("Supprimer """ & firstVal & """ ?", "Confirmation",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim dtRow As DataRow = Nothing
        For Each r As DataRow In dt.Rows
            If r(0).ToString() = firstVal Then
                dtRow = r
                Exit For
            End If
        Next
        If dtRow IsNot Nothing Then
            dt.Rows.Remove(dtRow)
            grid.DataSource = dt
            SetColWidths(grid)
            _modifie = True
            lblInfo.Text = "Fiche supprimee."
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' SAUVEGARDE
    ' ─────────────────────────────────────────────────────────────
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        SaveData()
    End Sub

    Public Sub SaveData()
        Try
            Dim dir As String = Path.GetDirectoryName(_xlsxPath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Using pkg As New ExcelPackage(New FileInfo(_xlsxPath))
                WriteSheet(pkg, SHEET_PHY, ColsPhy, DtPhy)
                WriteSheet(pkg, SHEET_MOR, ColsMor, DtMor)
                pkg.Save()
            End Using
            _modifie = False
            lblInfo.Text = "Sauvegarde : " & _xlsxPath
            MessageBox.Show("Fichier sauvegarde avec succes !", "OK",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Erreur sauvegarde : " & ex.Message, "Erreur",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub WriteSheet(pkg As ExcelPackage, sheetName As String, cols As String(), dt As DataTable)
        Dim existing = pkg.Workbook.Worksheets(sheetName)
        If existing IsNot Nothing Then pkg.Workbook.Worksheets.Delete(existing)
        Dim ws = pkg.Workbook.Worksheets.Add(sheetName)
        For c = 0 To cols.Length - 1
            ws.Cells(1, c + 1).Value = cols(c)
            ws.Cells(1, c + 1).Style.Font.Bold = True
        Next
        For r = 0 To dt.Rows.Count - 1
            For c = 0 To cols.Length - 1
                ws.Cells(r + 2, c + 1).Value = dt.Rows(r)(c).ToString()
            Next
        Next
    End Sub


    ' ─────────────────────────────────────────────────────────────
    ' FERMETURE
    ' ─────────────────────────────────────────────────────────────
    Private Sub PersonnesForm_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs)
        If _modifie Then
            Dim rep = MessageBox.Show("Modifications non sauvegardees. Sauvegarder ?",
                                      "Sauvegarder ?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If rep = DialogResult.Yes Then
                SaveData()
            ElseIf rep = DialogResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' HELPER UI
    ' ─────────────────────────────────────────────────────────────
    Private Function MkBtn(text As String, x As Integer, y As Integer,
                           w As Integer, bc As Color) As Button
        Dim b As New Button()
        b.Text = text
        b.Location = New Point(x, y)
        b.Size = New Size(w, 26)
        b.BackColor = bc
        b.ForeColor = Color.White
        b.FlatStyle = FlatStyle.Flat
        Return b
    End Function

End Class

' ═════════════════════════════════════════════════════════════════════════════
' FORMULAIRE DE SAISIE D'UNE FICHE
' ═════════════════════════════════════════════════════════════════════════════
Public Class FichePersonneForm
    Inherits Form

    Private _cols() As String
    Private _textFields() As TextBox
    Private _comboFields() As ComboBox
    Private _isCombo() As Boolean

    Public Function GetValues() As String()
        Dim result(_cols.Length - 1) As String
        For i = 0 To _cols.Length - 1
            If _isCombo(i) Then
                result(i) = _comboFields(i).Text.Trim()
            ElseIf _textFields(i) IsNot Nothing Then
                ' Si c'est le champ Editeur, récupérer la valeur complète depuis Tag
                If _textFields(i).Tag IsNot Nothing AndAlso
                   Not _textFields(i).Tag.ToString().StartsWith("editeur_main") Then
                    result(i) = _textFields(i).Tag.ToString()
                Else
                    result(i) = _textFields(i).Text.Trim()
                End If
            Else
                result(i) = ""
            End If
        Next
        Return result
    End Function

    Public Sub New(cols As String(), currentValues As String(), title As String)
        _cols = cols
        ' Sécuriser : remplacer tous les Nothing par ""
        If currentValues IsNot Nothing Then
            For i = 0 To currentValues.Length - 1
                If currentValues(i) Is Nothing Then currentValues(i) = ""
            Next
        End If
        Me.Text = title
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim rowH As Integer = 32
        Dim lblW As Integer = 170
        Dim fldW As Integer = 320
        Dim mg As Integer = 10
        Dim visibleRows As Integer = Math.Min(cols.Length, 18)
        Dim pnlH As Integer = visibleRows * rowH + mg
        Me.Size = New Size(lblW + fldW + mg * 4, pnlH + 70)

        Dim pnl As New Panel()
        pnl.AutoScroll = True
        pnl.Location = New Point(0, 0)
        pnl.Size = New Size(lblW + fldW + mg * 3, pnlH)
        pnl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        Me.Controls.Add(pnl)

        ReDim _textFields(cols.Length - 1)
        ReDim _comboFields(cols.Length - 1)
        ReDim _isCombo(cols.Length - 1)

        Dim dropdowns As New Dictionary(Of String, String())()
        dropdowns("Genre") = New String() {"MR", "MME", "MX"}
        dropdowns("Role") = New String() {"A", "C", "AR", "AD", "E", "A;C"}
        dropdowns("Rôle") = New String() {"A", "C", "AR", "AD", "E", "A;C"}
        dropdowns("SocieteGestion") = New String() {"SACEM", "SPEDIDAM", "ADAMI", "SCPP"}
        dropdowns("Forme Juridique") = New String() {"SAS", "SARL", "SA", "EURL", "EI", "SNC", "Association", "VAT", "LTD", "Autre"}
        dropdowns("Type de voie") = New String() {"Rue", "Avenue", "Boulevard", "Allée", "Impasse", "Place", "Chemin", "Route", "Passage", "Villa", "Cité", "Floor"}

        ' Index de la colonne Editeur (pour traitement spécial)
        Dim editeurIndex As Integer = -1
        For i = 0 To cols.Length - 1
            If cols(i).ToLower() = "editeur" Then editeurIndex = i : Exit For
        Next

        For i = 0 To cols.Length - 1
            Dim y As Integer = mg + i * rowH

            Dim lbl As New Label()
            lbl.Text = cols(i) & " :"
            lbl.Location = New Point(mg, y + 5)
            lbl.Size = New Size(lblW, 20)
            lbl.TextAlign = ContentAlignment.MiddleRight
            pnl.Controls.Add(lbl)

            Dim colKey As String = cols(i)

            ' Champ Id : lecture seule
            If colKey.ToLower() = "id" Then
                _isCombo(i) = False
                Dim tbId As New TextBox()
                tbId.Location = New Point(mg + lblW + mg, y + 2)
                tbId.Size = New Size(fldW, 23)
                tbId.ReadOnly = True
                tbId.BackColor = Color.FromArgb(235, 235, 235)
                If currentValues IsNot Nothing AndAlso i < currentValues.Length Then tbId.Text = currentValues(i)
                pnl.Controls.Add(tbId)
                _textFields(i) = tbId
                Continue For
            End If

            ' Champ Editeur : traitement spécial (liste multi-valeurs)
            If i = editeurIndex Then
                _isCombo(i) = False
                Dim tb As New TextBox()
                tb.Location = New Point(mg + lblW + mg, y + 2)
                tb.Size = New Size(fldW - 60, 23)
                tb.Tag = "editeur_main"
                If currentValues IsNot Nothing AndAlso i < currentValues.Length AndAlso currentValues(i) IsNot Nothing Then
                    ' Afficher uniquement la 1ère valeur — le reste est géré par lstEditeurs
                    Dim parts() As String = currentValues(i).Split(";"c)
                    tb.Text = parts(0).Trim()
                End If
                pnl.Controls.Add(tb)
                _textFields(i) = tb

                ' Bouton "..." pour ouvrir l'éditeur de liste
                Dim btnList As New Button()
                btnList.Text = "..."
                btnList.Location = New Point(mg + lblW + mg + fldW - 55, y + 1)
                btnList.Size = New Size(55, 25)
                btnList.Tag = i
                pnl.Controls.Add(btnList)

                ' Valeur complète stockée en Tag du TextBox
                If currentValues IsNot Nothing AndAlso i < currentValues.Length AndAlso currentValues(i) IsNot Nothing Then
                    tb.Tag = currentValues(i) ' stocker la valeur complète
                End If

                AddHandler btnList.Click, Sub(s, ev)
                                              Dim idx As Integer = CInt(DirectCast(s, Button).Tag)
                                              Dim currentVal As String = If(_textFields(idx).Tag IsNot Nothing,
                                                                             _textFields(idx).Tag.ToString(), "")
                                              Using dlg As New EditeurListForm(currentVal)
                                                  If dlg.ShowDialog() = DialogResult.OK Then
                                                      _textFields(idx).Tag = dlg.ValeurEditeurs
                                                      Dim first() As String = dlg.ValeurEditeurs.Split(";"c)
                                                      _textFields(idx).Text = If(first.Length > 0, first(0).Trim(), "")
                                                      If first.Length > 1 Then
                                                          _textFields(idx).Text &= " (+" & (first.Length - 1) & ")"
                                                      End If
                                                  End If
                                              End Using
                                          End Sub
                Continue For
            End If
            If dropdowns.ContainsKey(colKey) Then
                _isCombo(i) = True
                Dim cb As New ComboBox()
                cb.Location = New Point(mg + lblW + mg, y + 2)
                cb.Size = New Size(fldW, 23)
                cb.Items.AddRange(dropdowns(colKey))
                If currentValues IsNot Nothing AndAlso i < currentValues.Length Then
                    cb.Text = currentValues(i)
                End If
                pnl.Controls.Add(cb)
                _comboFields(i) = cb
            Else
                _isCombo(i) = False
                Dim tb As New TextBox()
                tb.Location = New Point(mg + lblW + mg, y + 2)
                tb.Size = New Size(fldW, 23)
                If currentValues IsNot Nothing AndAlso i < currentValues.Length Then
                    tb.Text = currentValues(i)
                End If
                pnl.Controls.Add(tb)
                _textFields(i) = tb
            End If
        Next

        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(90, 28)
        btnOK.Location = New Point(mg, pnlH + 5)
        btnOK.BackColor = Color.FromArgb(0, 120, 212)
        btnOK.ForeColor = Color.White
        btnOK.FlatStyle = FlatStyle.Flat
        btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Me.Controls.Add(btnOK)

        Dim btnCancel As New Button()
        btnCancel.Text = "Annuler"
        btnCancel.Size = New Size(90, 28)
        btnCancel.Location = New Point(mg + 100, pnlH + 5)
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Me.Controls.Add(btnCancel)
        Me.CancelButton = btnCancel

        ' Bouton Pappers uniquement pour les morales
        If title.ToLower().Contains("morale") Then
            Dim btnPappers As New Button()
            btnPappers.Text = "🔍 Rechercher sur Pappers"
            btnPappers.Size = New Size(200, 28)
            btnPappers.Location = New Point(mg + 210, pnlH + 5)
            btnPappers.BackColor = Color.FromArgb(30, 80, 160)
            btnPappers.ForeColor = Color.White
            btnPappers.FlatStyle = FlatStyle.Flat
            btnPappers.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            Me.Controls.Add(btnPappers)
            AddHandler btnPappers.Click, Sub(s, ev)
                                             Using dlg As New PappersSearchForm()
                                                 If dlg.ShowDialog() = DialogResult.OK Then
                                                     RemplirDepuisPappers(dlg.ResultatPappers)
                                                 End If
                                             End Using
                                         End Sub
        End If

        AddHandler btnOK.Click, AddressOf BtnOK_Click
    End Sub

    ''' <summary>Remplit les champs de la fiche depuis les données Pappers.</summary>
    Private Sub RemplirDepuisPappers(data As PappersData)
        For i = 0 To _cols.Length - 1
            Dim col As String = _cols(i).Trim()
            Dim val As String = ""
            ' Comparaison insensible casse ET accents
            Dim eq = Function(a As String) String.Compare(col, a, StringComparison.OrdinalIgnoreCase) = 0 OrElse
                                          String.Compare(col, a, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.CompareOptions.IgnoreCase Or System.Globalization.CompareOptions.IgnoreNonSpace) = 0
            If eq("Designation") Then : val = data.Denomination
            ElseIf eq("Siren") Then : val = data.Siren
            ElseIf eq("Forme Juridique") Then : val = data.FormeJuridique
            ElseIf eq("Capital") Then : val = data.Capital
            ElseIf eq("RCS") Then : val = data.RCS
            ElseIf eq("Num de voie") Then : val = data.NumVoie
            ElseIf eq("Type de voie") Then : val = data.TypeVoie
            ElseIf eq("Nom de voie") Then : val = data.NomVoie
            ElseIf eq("CP") Then : val = data.CP
            ElseIf eq("Ville") Then : val = data.Ville
            ElseIf eq("Pays") Then : val = data.Pays
            ElseIf eq("Prénom representant") Then : val = data.PrenomRepresentant
            ElseIf eq("Nom representant") Then : val = data.NomRepresentant
            ElseIf eq("Fonction representant") Then : val = data.FonctionRepresentant
            End If
            If Not String.IsNullOrEmpty(val) Then
                If _isCombo(i) AndAlso _comboFields(i) IsNot Nothing Then
                    _comboFields(i).Text = val
                ElseIf Not _isCombo(i) AndAlso _textFields(i) IsNot Nothing Then
                    _textFields(i).Text = val
                End If
            End If
        Next
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        Dim firstVal As String = If(_isCombo(0), _comboFields(0).Text.Trim(), _textFields(0).Text.Trim())
        If String.IsNullOrEmpty(firstVal) Then
            MessageBox.Show("Le champ """ & _cols(0) & """ est obligatoire.",
                            "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class

' ═════════════════════════════════════════════════════════════════════════════
' FORMULAIRE LISTE DES ÉDITEURS (multi-valeurs séparées par ;)
' ═════════════════════════════════════════════════════════════════════════════
Public Class EditeurListForm
    Inherits Form

    Public ReadOnly Property ValeurEditeurs As String
        Get
            Dim items As New List(Of String)()
            For Each item As Object In lstEditeurs.Items
                If Not String.IsNullOrEmpty(item.ToString().Trim()) Then
                    items.Add(item.ToString().Trim())
                End If
            Next
            Return String.Join(";", items)
        End Get
    End Property

    Private lstEditeurs As ListBox
    Private txtNouvel As TextBox
    Private _dtMor As DataTable

    Public Sub New(currentValue As String)
        Me.Text = "Éditeurs associés"
        Me.Size = New Size(450, 380)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' Charger les morales pour la ComboBox
        _dtMor = New DataTable()
        Try
            If File.Exists(PersonnesForm.DefaultXlsxPath) Then
                Using pkg As New ExcelPackage(New FileInfo(PersonnesForm.DefaultXlsxPath))
                    Dim ws = pkg.Workbook.Worksheets("PERSONNEMORALE")
                    If ws IsNot Nothing AndAlso ws.Dimension IsNot Nothing Then
                        For c = 1 To ws.Dimension.Columns
                            _dtMor.Columns.Add(ws.Cells(1, c).Text.Trim())
                        Next
                        For r = 2 To ws.Dimension.Rows
                            Dim nr As DataRow = _dtMor.NewRow()
                            For c = 1 To ws.Dimension.Columns
                                nr(c - 1) = ws.Cells(r, c).Text.Trim()
                            Next
                            _dtMor.Rows.Add(nr)
                        Next
                    End If
                End Using
            End If
        Catch
        End Try

        ' Liste des éditeurs actuels
        Dim lblListe As New Label()
        lblListe.Text = "Éditeurs associés :"
        lblListe.Location = New Point(10, 10)
        lblListe.AutoSize = True
        Me.Controls.Add(lblListe)

        lstEditeurs = New ListBox()
        lstEditeurs.Location = New Point(10, 30)
        lstEditeurs.Size = New Size(410, 180)
        Me.Controls.Add(lstEditeurs)

        ' Peupler avec valeurs existantes
        If Not String.IsNullOrEmpty(currentValue) Then
            For Each v As String In currentValue.Split(";"c)
                Dim trimmed As String = v.Trim()
                If Not String.IsNullOrEmpty(trimmed) Then lstEditeurs.Items.Add(trimmed)
            Next
        End If
        ' Ajouter EAC comme option spéciale
        If Not lstEditeurs.Items.Contains("EAC") Then lstEditeurs.Items.Insert(0, "EAC")

        ' Zone d'ajout
        Dim lblAjouter As New Label()
        lblAjouter.Text = "Ajouter un éditeur :"
        lblAjouter.Location = New Point(10, 220)
        lblAjouter.AutoSize = True
        Me.Controls.Add(lblAjouter)

        ' ComboBox alimentée par les morales
        Dim cbNouvel As New ComboBox()
        cbNouvel.Location = New Point(10, 240)
        cbNouvel.Size = New Size(310, 23)
        cbNouvel.DropDownStyle = ComboBoxStyle.DropDown
        cbNouvel.Items.Add("EAC")
        For Each row As DataRow In _dtMor.Rows
            Dim desig As String = row(0).ToString().Trim()
            If Not String.IsNullOrEmpty(desig) Then cbNouvel.Items.Add(desig)
        Next
        Me.Controls.Add(cbNouvel)

        Dim btnAjouter As New Button()
        btnAjouter.Text = "+ Ajouter"
        btnAjouter.Location = New Point(330, 239)
        btnAjouter.Size = New Size(90, 25)
        btnAjouter.BackColor = Color.FromArgb(16, 124, 16)
        btnAjouter.ForeColor = Color.White
        btnAjouter.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnAjouter)

        Dim btnSupprimer As New Button()
        btnSupprimer.Text = "Supprimer"
        btnSupprimer.Location = New Point(10, 275)
        btnSupprimer.Size = New Size(90, 25)
        btnSupprimer.BackColor = Color.FromArgb(196, 43, 28)
        btnSupprimer.ForeColor = Color.White
        btnSupprimer.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnSupprimer)

        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Location = New Point(240, 310)
        btnOK.Size = New Size(85, 28)
        btnOK.BackColor = Color.FromArgb(0, 120, 212)
        btnOK.ForeColor = Color.White
        btnOK.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnOK)

        Dim btnCancel As New Button()
        btnCancel.Text = "Annuler"
        btnCancel.Location = New Point(335, 310)
        btnCancel.Size = New Size(85, 28)
        btnCancel.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnCancel)
        Me.CancelButton = btnCancel

        AddHandler btnAjouter.Click, Sub(s, ev)
                                         Dim val As String = cbNouvel.Text.Trim()
                                         If Not String.IsNullOrEmpty(val) AndAlso Not lstEditeurs.Items.Contains(val) Then
                                             lstEditeurs.Items.Add(val)
                                             cbNouvel.Text = ""
                                         End If
                                     End Sub

        AddHandler btnSupprimer.Click, Sub(s, ev)
                                           If lstEditeurs.SelectedIndex >= 0 Then
                                               lstEditeurs.Items.RemoveAt(lstEditeurs.SelectedIndex)
                                           End If
                                       End Sub

        AddHandler btnOK.Click, Sub(s, ev)
                                    Me.DialogResult = DialogResult.OK
                                    Me.Close()
                                End Sub
    End Sub

End Class

' ═════════════════════════════════════════════════════════════════════════════
' MODÈLE DE DONNÉES PAPPERS
' ═════════════════════════════════════════════════════════════════════════════
Public Class PappersData
    Public Property Denomination As String = ""
    Public Property Siren As String = ""
    Public Property FormeJuridique As String = ""
    Public Property Capital As String = ""
    Public Property RCS As String = ""
    Public Property NumVoie As String = ""
    Public Property TypeVoie As String = ""
    Public Property NomVoie As String = ""
    Public Property CP As String = ""
    Public Property Ville As String = ""
    Public Property Pays As String = ""
    Public Property PrenomRepresentant As String = ""
    Public Property NomRepresentant As String = ""
    Public Property FonctionRepresentant As String = ""
End Class

' ═════════════════════════════════════════════════════════════════════════════
' FORMULAIRE RECHERCHE PAPPERS (scraping)
' ═════════════════════════════════════════════════════════════════════════════
Public Class PappersSearchForm
    Inherits Form

    Private _ResultatPappers As PappersData
    Public ReadOnly Property ResultatPappers As PappersData
        Get
            Return _ResultatPappers
        End Get
    End Property

    Private txtRecherche As TextBox
    Private lstResultats As ListBox
    Private lblStatut As Label
    Private btnRechercher As Button
    Private _resultats As New List(Of PappersData)()

    Public Sub New()
        Me.Text = "Rechercher une entreprise sur Pappers"
        Me.Size = New Size(560, 420)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblRecherche As New Label()
        lblRecherche.Text = "SIREN ou nom de l'entreprise :"
        lblRecherche.Location = New Point(10, 15)
        lblRecherche.AutoSize = True
        Me.Controls.Add(lblRecherche)

        txtRecherche = New TextBox()
        txtRecherche.Location = New Point(10, 35)
        txtRecherche.Size = New Size(390, 23)
        Me.Controls.Add(txtRecherche)

        btnRechercher = New Button()
        btnRechercher.Text = "Rechercher"
        btnRechercher.Location = New Point(410, 34)
        btnRechercher.Size = New Size(120, 25)
        btnRechercher.BackColor = Color.FromArgb(30, 80, 160)
        btnRechercher.ForeColor = Color.White
        btnRechercher.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnRechercher)

        lblStatut = New Label()
        lblStatut.Location = New Point(10, 65)
        lblStatut.Size = New Size(520, 20)
        lblStatut.ForeColor = Color.Gray
        lblStatut.Text = "Saisissez un SIREN (9 chiffres) ou un nom, puis cliquez Rechercher."
        Me.Controls.Add(lblStatut)

        Dim lblRes As New Label()
        lblRes.Text = "Résultats :"
        lblRes.Location = New Point(10, 90)
        lblRes.AutoSize = True
        Me.Controls.Add(lblRes)

        lstResultats = New ListBox()
        lstResultats.Location = New Point(10, 110)
        lstResultats.Size = New Size(520, 200)
        Me.Controls.Add(lstResultats)

        Dim btnSelectionner As New Button()
        btnSelectionner.Text = "Sélectionner"
        btnSelectionner.Location = New Point(330, 325)
        btnSelectionner.Size = New Size(100, 28)
        btnSelectionner.BackColor = Color.FromArgb(0, 120, 212)
        btnSelectionner.ForeColor = Color.White
        btnSelectionner.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnSelectionner)

        Dim btnAnnuler As New Button()
        btnAnnuler.Text = "Annuler"
        btnAnnuler.Location = New Point(440, 325)
        btnAnnuler.Size = New Size(90, 28)
        btnAnnuler.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnAnnuler)
        Me.CancelButton = btnAnnuler

        AddHandler btnRechercher.Click, AddressOf Rechercher
        AddHandler txtRecherche.KeyDown, Sub(s, ev)
                                             If ev.KeyCode = Keys.Return Then Rechercher(s, ev)
                                         End Sub
        AddHandler btnSelectionner.Click, Sub(s, ev)
                                              If lstResultats.SelectedIndex >= 0 Then
                                                  EnrichirEtValider(lstResultats.SelectedIndex)
                                              Else
                                                  MessageBox.Show("Sélectionnez un résultat.", "Attention",
                                                                  MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                              End If
                                          End Sub
        AddHandler lstResultats.DoubleClick, Sub(s, ev)
                                                 If lstResultats.SelectedIndex >= 0 Then
                                                     EnrichirEtValider(lstResultats.SelectedIndex)
                                                 End If
                                             End Sub
    End Sub

    ''' <summary>
    ''' Appelé quand l'utilisateur sélectionne un résultat.
    ''' Enrichit les données via societe.com (capital, dirigeant, adresse complète)
    ''' puis ferme avec OK.
    ''' </summary>
    Private Sub EnrichirEtValider(idx As Integer)
        Dim base As PappersData = _resultats(idx)
        If Not String.IsNullOrEmpty(base.Siren) Then
            Try
                Dim ficheUrl As String = "https://www.societe.com/societe/a-" & base.Siren & ".html"
                Dim html As String = AppelSociete(ficheUrl)
                If Not String.IsNullOrEmpty(html) Then
                    ' Capital
                    Dim capital As String = ExtraireCapitalSociete(html)
                    If Not String.IsNullOrEmpty(capital) Then base.Capital = capital
                    ' Forme juridique
                    Dim formeRaw As String = HtmlVal(html, "data-copy-id=""resume_legal_label"">", "</template>").Trim()
                    If Not String.IsNullOrEmpty(formeRaw) Then
                        base.FormeJuridique = MapFormeJuridique(formeRaw)
                    End If
                End If
            Catch ex As Exception
                ' Pas grave si societe.com échoue
            End Try
        End If
        _ResultatPappers = base
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Rechercher(sender As Object, e As EventArgs)
        Dim query As String = txtRecherche.Text.Trim()
        If String.IsNullOrEmpty(query) Then Return
        lblStatut.Text = "Recherche en cours..."
        btnRechercher.Enabled = False
        lstResultats.Items.Clear()
        _resultats.Clear()
        Application.DoEvents()
        Try
            Dim isSiren As Boolean = (query.Length = 9 AndAlso query.All(Function(c) Char.IsDigit(c)))
            If isSiren Then
                Dim data As PappersData = AppelerApiFiche(query)
                If data IsNot Nothing Then
                    _resultats.Add(data)
                    lstResultats.Items.Add(FormatResultat(data))
                    lblStatut.Text = "1 résultat trouvé."
                Else
                    lblStatut.Text = "Aucune entreprise trouvée pour ce SIREN."
                End If
            Else
                AppelerApiRecherche(query, _resultats)
                For Each r As PappersData In _resultats
                    lstResultats.Items.Add(FormatResultat(r))
                Next
                lblStatut.Text = _resultats.Count & " résultat(s) trouvé(s)."
            End If
        Catch ex As Exception
            lblStatut.Text = "Erreur : " & ex.Message
        Finally
            btnRechercher.Enabled = True
        End Try
    End Sub

    Private Function FormatResultat(d As PappersData) As String
        Return d.Denomination & " — " & d.Siren & " — " & d.FormeJuridique & " — " & d.Ville
    End Function

    Private Function AppelApi(url As String) As String
        Using client As New System.Net.Http.HttpClient()
            client.Timeout = TimeSpan.FromSeconds(10)
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0")
            client.DefaultRequestHeaders.Add("Accept", "application/json")
            Dim response = client.GetAsync(url).Result
            If Not response.IsSuccessStatusCode Then
                Throw New Exception("HTTP " & CInt(response.StatusCode) & " - " & response.ReasonPhrase)
            End If
            Return response.Content.ReadAsStringAsync().Result
        End Using
    End Function

    Private Const PAPPERS_KEY As String = "7908be8000ea642aaabedeb73f3100bc57240b71ab4b5c4"

    Private Function AppelerApiFiche(siren As String) As PappersData
        ' Fiche via API gouv recherche-entreprises (SIREN direct)
        Dim url As String = "https://recherche-entreprises.api.gouv.fr/search?q=" & siren & "&nombre=1"
        Dim json As String = AppelJson(url)
        Dim bStart As Integer = json.IndexOf("{""siren""")
        If bStart < 0 Then bStart = json.IndexOf("{")
        If bStart < 0 Then Return Nothing
        ' Extraire le premier bloc résultat
        Dim depth As Integer = 0
        For i As Integer = bStart To json.Length - 1
            If json(i) = "{"c Then
                depth += 1
            ElseIf json(i) = "}"c Then
                depth -= 1
                If depth = 0 Then
                    Return ParseGouvBlock(json.Substring(bStart, i - bStart + 1))
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Sub AppelerApiRecherche(query As String, resultats As List(Of PappersData))
        Dim url As String = "https://recherche-entreprises.api.gouv.fr/search?q=" &
                            Uri.EscapeDataString(query) & "&nombre=10"
        Dim json As String = AppelJson(url)
        Dim arrIdx As Integer = json.IndexOf("""results""")
        If arrIdx < 0 Then Return
        Dim arrStart As Integer = json.IndexOf("[", arrIdx)
        If arrStart < 0 Then Return
        Dim depth As Integer = 0
        Dim bStart As Integer = -1
        For i As Integer = arrStart To json.Length - 1
            If json(i) = "{"c Then
                If depth = 0 Then bStart = i
                depth += 1
            ElseIf json(i) = "}"c Then
                depth -= 1
                If depth = 0 AndAlso bStart >= 0 Then
                    Dim d As PappersData = ParseGouvBlock(json.Substring(bStart, i - bStart + 1))
                    If d IsNot Nothing AndAlso Not String.IsNullOrEmpty(d.Denomination) Then
                        resultats.Add(d)
                        If resultats.Count >= 10 Then Return
                    End If
                    bStart = -1
                End If
            ElseIf json(i) = "]"c AndAlso depth = 0 Then
                Exit For
            End If
        Next
    End Sub

    ''' <summary>Parse un bloc JSON résultat de recherche-entreprises.api.gouv.fr</summary>
    Private Function ParseGouvBlock(block As String) As PappersData
        Dim d As New PappersData()
        d.Denomination = JsonVal(block, "nom_complet")
        If String.IsNullOrEmpty(d.Denomination) Then d.Denomination = JsonVal(block, "nom_raison_sociale")
        d.Siren = JsonVal(block, "siren")
        Dim formeLib As String = JsonVal(block, "libelle_nature_juridique_principale")
        d.FormeJuridique = If(Not String.IsNullOrEmpty(formeLib), MapFormeJuridique(formeLib), "")
        d.Pays = "France"

        ' Siège → adresse
        Dim siegeIdx As Integer = block.IndexOf("""siege""")
        If siegeIdx >= 0 Then
            Dim siegeBlock As String = ExtractJsonBlock(block, siegeIdx)
            d.NumVoie = JsonVal(siegeBlock, "numero_voie")
            ' TypeVoie : mapper vers valeurs dropdown
            Dim tvRaw As String = JsonVal(siegeBlock, "type_voie").ToUpper().Trim()
            Dim typeMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"RUE", "Rue"}, {"R", "Rue"}, {"AV", "Avenue"}, {"AVE", "Avenue"}, {"AVENUE", "Avenue"},
                {"BD", "Boulevard"}, {"BLD", "Boulevard"}, {"BOULEVARD", "Boulevard"},
                {"ALL", "Allée"}, {"ALLEE", "Allée"}, {"ALLÉE", "Allée"},
                {"IMP", "Impasse"}, {"IMPASSE", "Impasse"},
                {"PL", "Place"}, {"PLACE", "Place"},
                {"CHE", "Chemin"}, {"CH", "Chemin"}, {"CHEMIN", "Chemin"},
                {"RTE", "Route"}, {"ROUTE", "Route"},
                {"PAS", "Passage"}, {"PASSAGE", "Passage"},
                {"VLA", "Villa"}, {"VILLA", "Villa"},
                {"SQ", "Square"}, {"SQUARE", "Square"}}
            If typeMap.ContainsKey(tvRaw) Then
                d.TypeVoie = typeMap(tvRaw)
            ElseIf Not String.IsNullOrEmpty(tvRaw) Then
                d.TypeVoie = ToTitleCase(tvRaw)
            End If
            d.NomVoie = ToTitleCase(JsonVal(siegeBlock, "libelle_voie"))
            d.CP = JsonVal(siegeBlock, "code_postal")
            d.Ville = ToTitleCase(JsonVal(siegeBlock, "libelle_commune"))
            d.RCS = ToTitleCase(JsonVal(siegeBlock, "libelle_commune"))
        End If

        ' Capital
        Dim capVal As String = ""
        Dim complIdx As Integer = block.IndexOf("""complements""")
        If complIdx >= 0 Then
            Dim complBlock As String = ExtractJsonBlock(block, complIdx)
            capVal = JsonVal(complBlock, "capital_social")
        End If
        ' Fallback : chercher capital_social directement dans tout le block
        If String.IsNullOrEmpty(capVal) OrElse capVal = "0" OrElse capVal = "null" Then
            capVal = JsonVal(block, "capital_social")
        End If
        If String.IsNullOrEmpty(capVal) OrElse capVal = "0" OrElse capVal = "null" Then
            Dim finIdx As Integer = block.IndexOf("""finances""")
            If finIdx >= 0 Then
                Dim finBlock As String = ExtractJsonBlock(block, finIdx)
                capVal = JsonVal(finBlock, "capital")
            End If
        End If
        If String.IsNullOrEmpty(capVal) OrElse capVal = "null" Then capVal = JsonVal(block, "capital")
        If Not String.IsNullOrEmpty(capVal) AndAlso capVal <> "null" AndAlso capVal <> "0" Then
            ' Formater : 100 → "100 €", 100000 → "100 000 €"
            Dim capNum As Decimal
            If Decimal.TryParse(capVal, capNum) Then
                d.Capital = String.Format("{0:N0}", capNum).Replace(",", " ") & " €"
            Else
                d.Capital = capVal & " €"
            End If
        End If

        ' Dirigeants
        Dim dirgIdx As Integer = block.IndexOf("""dirigeants""")
        If dirgIdx >= 0 Then
            Dim arrStart As Integer = block.IndexOf("[", dirgIdx)
            If arrStart >= 0 Then
                Dim depth As Integer = 0
                Dim bStart As Integer = -1
                For i As Integer = arrStart To block.Length - 1
                    If block(i) = "{"c Then
                        If depth = 0 Then bStart = i
                        depth += 1
                    ElseIf block(i) = "}"c Then
                        depth -= 1
                        If depth = 0 AndAlso bStart >= 0 Then
                            Dim pp As String = block.Substring(bStart, i - bStart + 1)
                            If JsonVal(pp, "type_dirigeant") <> "personne morale" Then
                                Dim prenomsBrut As String = JsonVal(pp, "prenoms")
                                If String.IsNullOrEmpty(prenomsBrut) Then prenomsBrut = JsonVal(pp, "prenom")
                                d.PrenomRepresentant = ToTitleCase(prenomsBrut)
                                d.NomRepresentant = JsonVal(pp, "nom").ToUpper()
                                d.FonctionRepresentant = PremierMot(JsonVal(pp, "qualite"))
                                Exit For
                            End If
                            bStart = -1
                        End If
                    ElseIf block(i) = "]"c AndAlso depth = 0 Then
                        Exit For
                    End If
                Next
            End If
        End If
        Return d
    End Function

    ''' <summary>Retourne le premier mot d'une chaîne (ex: "Président de SAS" → "Président").</summary>
    Private Function PremierMot(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim idx As Integer = s.IndexOf(" ")
        If idx > 0 Then Return s.Substring(0, idx)
        Return s
    End Function

    ''' <summary>Extrait uniquement le capital depuis le HTML societe.com.</summary>
    Private Function ExtraireCapitalSociete(html As String) As String
        If String.IsNullOrEmpty(html) Then Return ""
        Dim capIdx As Integer = html.IndexOf("Capital : ")
        If capIdx >= 0 Then
            Dim capStart As Integer = capIdx + 10
            Dim capEnd As Integer = capStart
            Do While capEnd < html.Length AndAlso html(capEnd) <> "."c AndAlso
                     html(capEnd) <> "<"c AndAlso html(capEnd) <> ";"c
                capEnd += 1
            Loop
            Dim capVal As String = html.Substring(capStart, capEnd - capStart).Trim()
            If Not String.IsNullOrEmpty(capVal) AndAlso capVal <> "0" Then
                Dim capNum As Decimal
                If Decimal.TryParse(capVal, capNum) Then
                    Return String.Format("{0:N0}", capNum).Replace(",", " ") & " €"
                End If
                Return capVal & " €"
            End If
        End If
        Return ""
    End Function

    Private Function AppelSociete(url As String) As String
        Dim wc As New System.Net.WebClient()
        wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36")
        wc.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
        wc.Headers.Add("Accept-Language", "fr-FR,fr;q=0.9")
        wc.Encoding = System.Text.Encoding.UTF8
        Return wc.DownloadString(url)
    End Function

    Private Function AppelJson(url As String) As String
        Dim wc As New System.Net.WebClient()
        wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        wc.Headers.Add("Accept", "application/json")
        wc.Encoding = System.Text.Encoding.UTF8
        Return wc.DownloadString(url)
    End Function

    ''' <summary>Parse la fiche HTML societe.com (résultat direct SIREN).</summary>
    Private Function ParseSocieteHtml(html As String) As PappersData
        If String.IsNullOrEmpty(html) Then Return Nothing
        Dim d As New PappersData()

        ' === NOM : <h1>LONGUE VIE PUBLISHING</h1> ===
        d.Denomination = HtmlVal(html, "<h1>", "</h1>").Trim()

        ' === SIREN : data-sid="940571730" ===
        d.Siren = HtmlAttr(html, "data-sid=""", """")

        ' === FORME JURIDIQUE : <template data-copy-id="resume_legal_label">Société par actions simplifiée</template> ===
        Dim formeRaw As String = HtmlVal(html, "data-copy-id=""resume_legal_label"">", "</template>").Trim()
        d.FormeJuridique = MapFormeJuridique(formeRaw)

        ' === ADRESSE : <template data-copy-id="resume_company_address">61 RUE DE LYON, 75012 PARIS</template> ===
        Dim adresseRaw As String = HtmlVal(html, "data-copy-id=""resume_company_address"">", "</template>").Trim()
        If Not String.IsNullOrEmpty(adresseRaw) Then ParseAdresseSociete(adresseRaw, d)

        ' === CAPITAL : "Capital : 100 ." → "100 €"
        Dim capIdx As Integer = html.IndexOf("Capital : ")
        If capIdx >= 0 Then
            Dim capStart As Integer = capIdx + 10
            Dim capEnd As Integer = capStart
            Do While capEnd < html.Length AndAlso html(capEnd) <> "."c AndAlso html(capEnd) <> "<"c AndAlso html(capEnd) <> ";"c
                capEnd += 1
            Loop
            Dim capVal As String = html.Substring(capStart, capEnd - capStart).Trim()
            If Not String.IsNullOrEmpty(capVal) AndAlso capVal <> "0" Then
                d.Capital = capVal & " €"
            End If
        End If

        ' === PAYS : societe.com = toujours France pour les entreprises françaises ===
        d.Pays = "France"

        ' === RCS : ville depuis l'adresse ===
        If Not String.IsNullOrEmpty(d.Ville) Then d.RCS = ToTitleCase(d.Ville)

        ' === DIRIGEANT : JSON-LD en priorité (fiable pour toutes les fiches societe.com) ===
        ' "givenName":"Landry","familyName":"AGONHOUMEY","jobTitle":"Pr\u00E9sident"
        Dim nomTrouve As Boolean = False
        Dim prenomJld As String = DecodeJsonUnicode(JsonValHtml(html, "givenName"))
        Dim nomJld As String = DecodeJsonUnicode(JsonValHtml(html, "familyName"))
        Dim fctJld As String = DecodeJsonUnicode(JsonValHtml(html, "jobTitle"))
        If Not String.IsNullOrEmpty(nomJld) Then
            d.PrenomRepresentant = ToTitleCase(prenomJld)
            d.NomRepresentant = nomJld.ToUpper()
            d.FonctionRepresentant = PremierMot(fctJld)
            nomTrouve = True
        End If
        ' Fallback : <span class="ui-label" [aria-hidden]?>Prénom NOM</span>
        ' Filtre strict : exige au moins un mot entièrement en MAJUSCULES = NOM
        If Not nomTrouve Then
            Dim exclusions() As String = {"SOCIAL", "JURIDIQUE", "CAPITAL", "SIEGE", "ACTIVITE",
                                           "EFFECTIF", "TVA", "SIRET", "APE", "NAF", "RCS",
                                           "connecter", "inscrire", "abonner"}
            Dim spanIdx2 As Integer = 0
            Do
                Dim sIdx As Integer = html.IndexOf("ui-label""", spanIdx2, StringComparison.OrdinalIgnoreCase)
                If sIdx < 0 Then Exit Do
                Dim tagClose As Integer = html.IndexOf(">", sIdx)
                If tagClose < 0 Then Exit Do
                Dim contentStart As Integer = tagClose + 1
                Dim contentEnd As Integer = html.IndexOf("</span>", contentStart, StringComparison.OrdinalIgnoreCase)
                If contentEnd < 0 Then Exit Do
                Dim fullName As String = System.Net.WebUtility.HtmlDecode(
                    html.Substring(contentStart, contentEnd - contentStart).Trim())
                If fullName.Contains("<") OrElse fullName.Contains("""") Then
                    spanIdx2 = contentEnd : Continue Do
                End If
                Dim estExclu As Boolean = exclusions.Any(Function(ex) fullName.ToUpper().Contains(ex.ToUpper()))
                If Not estExclu AndAlso fullName.Length > 2 Then
                    Dim parts() As String = fullName.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim nomsFb As New List(Of String)
                    Dim prenomsFb As New List(Of String)
                    For Each p As String In parts
                        If p.Length > 1 AndAlso p = p.ToUpper() AndAlso
                           p.All(Function(c) Char.IsLetter(c) OrElse c = "-"c OrElse c = "'"c) Then
                            nomsFb.Add(p)
                        Else
                            prenomsFb.Add(ToTitleCase(p))
                        End If
                    Next
                    If nomsFb.Count > 0 Then
                        d.NomRepresentant = String.Join(" ", nomsFb)
                        d.PrenomRepresentant = String.Join(" ", prenomsFb)
                        nomTrouve = True
                        Dim suite As String = html.Substring(contentEnd, Math.Min(400, html.Length - contentEnd))
                        For Each fct2 As String In {"Président directeur général",
                                                    "Président du conseil d'administration",
                                                    "Président", "Gérant", "Directeur général délégué",
                                                    "Directeur général", "Administrateur général",
                                                    "Administrateur délégué", "Administrateur",
                                                    "Co-gérant", "Co-président", "Vice-président",
                                                    "Associé gérant", "Associé unique", "Associé",
                                                    "Directeur", "Liquidateur", "Commissaire aux comptes"}
                            If suite.IndexOf(fct2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                d.FonctionRepresentant = PremierMot(fct2) : Exit For
                            End If
                        Next
                        Exit Do
                    End If
                End If
                spanIdx2 = contentEnd
            Loop
        End If

        Return d
    End Function

    ''' <summary>Décode les escapes Unicode JSON \uXXXX en caractères réels.</summary>
    Private Function DecodeJsonUnicode(s As String) As String
        If String.IsNullOrEmpty(s) Then Return s
        Dim sb As New System.Text.StringBuilder()
        Dim i As Integer = 0
        Do While i < s.Length
            If i + 5 < s.Length AndAlso s(i) = "\"c AndAlso s(i + 1) = "u"c Then
                Dim hex As String = s.Substring(i + 2, 4)
                Dim code As Integer
                If Integer.TryParse(hex, Globalization.NumberStyles.HexNumber, Nothing, code) Then
                    sb.Append(ChrW(code))
                    i += 6
                    Continue Do
                End If
            End If
            sb.Append(s(i))
            i += 1
        Loop
        Return sb.ToString()
    End Function

    Private Function MapFormeJuridique(formeRaw As String) As String
        Dim f As String = formeRaw.ToUpper()
        If f.Contains("ACTION") AndAlso f.Contains("SIMPLIFI") Then Return "SAS"
        If f.Contains("UNIPERSONNELLE") AndAlso f.Contains("RESPONSABILIT") Then Return "EURL"
        If f.Contains("RESPONSABILIT") AndAlso f.Contains("LIMIT") Then Return "SARL"
        If f.Contains("ANONYME") Then Return "SA"
        If f.Contains("NOM COLLECTIF") Then Return "SNC"
        If f.Contains("INDIVIDU") Then Return "EI"
        If f.Contains("ASSOCIATION") Then Return "Association"
        Dim parenIdx As Integer = formeRaw.LastIndexOf("(")
        If parenIdx >= 0 Then
            Dim parenEnd As Integer = formeRaw.IndexOf(")", parenIdx)
            If parenEnd > parenIdx Then
                Select Case formeRaw.Substring(parenIdx + 1, parenEnd - parenIdx - 1).Trim().ToUpper()
                    Case "SAS", "SASU" : Return "SAS"
                    Case "SARL" : Return "SARL"
                    Case "SA" : Return "SA"
                    Case "EURL" : Return "EURL"
                    Case "EI" : Return "EI"
                    Case "SNC" : Return "SNC"
                End Select
            End If
        End If
        Return "Autre"
    End Function

    ''' <summary>Parse la liste de résultats societe.com (recherche par nom).</summary>
    Private Sub ParseSocieteResultats(html As String, resultats As List(Of PappersData))
        ' Si redirection directe vers une fiche (SIREN unique) → parser comme fiche
        If html.Contains("data-sid=""") AndAlso Not html.Contains("class=""ui-card""") Then
            Dim d As PappersData = ParseSocieteHtml(html)
            If d IsNot Nothing AndAlso Not String.IsNullOrEmpty(d.Denomination) Then
                resultats.Add(d)
            End If
            Return
        End If

        ' Liste de résultats — chercher les blocs entreprise
        ' Pattern : href="/societe/nom-siren.html"
        Dim idx As Integer = 0
        Do
            Dim cardIdx As Integer = html.IndexOf("/societe/", idx)
            If cardIdx < 0 Then Exit Do
            Dim hrefEnd As Integer = html.IndexOf(".html", cardIdx)
            If hrefEnd < 0 Then Exit Do
            Dim slug As String = html.Substring(cardIdx + 9, hrefEnd - cardIdx - 9)
            ' Extraire SIREN (derniers chiffres après le dernier -)
            Dim lastDash As Integer = slug.LastIndexOf("-")
            If lastDash >= 0 Then
                Dim siren As String = slug.Substring(lastDash + 1)
                If siren.Length = 9 AndAlso siren.All(Function(c) Char.IsDigit(c)) AndAlso
                   Not resultats.Any(Function(r) r.Siren = siren) Then
                    ' Trouver nom dans le contexte
                    Dim ctxStart As Integer = Math.Max(0, cardIdx - 100)
                    Dim ctxEnd As Integer = Math.Min(html.Length, hrefEnd + 300)
                    Dim ctx As String = html.Substring(ctxStart, ctxEnd - ctxStart)
                    Dim nom As String = HtmlVal(ctx, "<h3>", "</h3>")
                    If String.IsNullOrEmpty(nom) Then nom = HtmlVal(ctx, "<h4>", "</h4>")
                    If String.IsNullOrEmpty(nom) Then nom = slug.Substring(0, lastDash).Replace("-", " ").ToUpper()
                    Dim ville As String = HtmlVal(ctx, "class=""ui-city"">", "</")
                    Dim d As New PappersData()
                    d.Siren = siren
                    d.Denomination = System.Net.WebUtility.HtmlDecode(nom).Trim()
                    d.Ville = ville.Trim()
                    If Not String.IsNullOrEmpty(d.Denomination) Then
                        resultats.Add(d)
                        If resultats.Count >= 10 Then Return
                    End If
                End If
            End If
            idx = hrefEnd + 5
        Loop
    End Sub

    Private Sub ParseAdresseSociete(adresse As String, d As PappersData)
        If String.IsNullOrEmpty(adresse) Then Return
        ' Format societe.com : "61 RUE DE LYON, 75012 PARIS" ou "61 BIS AVENUE DES FLEURS, 75012 PARIS"
        Dim commaIdx As Integer = adresse.IndexOf(",")
        If commaIdx < 0 Then Return
        Dim avant As String = adresse.Substring(0, commaIdx).Trim()
        Dim apres As String = adresse.Substring(commaIdx + 1).Trim()
        ' CP + Ville
        Dim mCP = System.Text.RegularExpressions.Regex.Match(apres, "^(\d{5})\s+(.+)$")
        If mCP.Success Then
            d.CP = mCP.Groups(1).Value
            d.Ville = ToTitleCase(mCP.Groups(2).Value.Trim())
        End If
        ' Mapping types de voie bruts → valeurs du dropdown
        Dim typeMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"RUE", "Rue"}, {"R", "Rue"},
            {"AVENUE", "Avenue"}, {"AVE", "Avenue"}, {"AV", "Avenue"},
            {"BOULEVARD", "Boulevard"}, {"BD", "Boulevard"}, {"BLD", "Boulevard"}, {"BLVD", "Boulevard"},
            {"ALLEE", "Allée"}, {"ALLÉE", "Allée"}, {"ALL", "Allée"},
            {"IMPASSE", "Impasse"}, {"IMP", "Impasse"},
            {"PLACE", "Place"}, {"PL", "Place"},
            {"CHEMIN", "Chemin"}, {"CHE", "Chemin"}, {"CH", "Chemin"},
            {"ROUTE", "Route"}, {"RTE", "Route"},
            {"PASSAGE", "Passage"}, {"PAS", "Passage"},
            {"VILLA", "Villa"}, {"VLA", "Villa"},
            {"CITE", "Cité"}, {"CITÉ", "Cité"},
            {"SQUARE", "Square"}, {"SQ", "Square"},
            {"VOIE", "Rue"}, {"QUAI", "Rue"}, {"RUELLE", "Rue"},
            {"LOTISSEMENT", "Allée"}, {"LOT", "Allée"},
            {"DOMAINE", "Allée"}, {"DOM", "Allée"},
            {"HAMEAU", "Allée"}, {"LIEU DIT", "Allée"}}
        Dim tokens() As String = avant.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        If tokens.Length >= 2 Then
            ' Numéro = premier token (peut être "61" ou "61BIS" ou "61 BIS")
            Dim numVoie As String = tokens(0)
            Dim startSearch As Integer = 1
            ' Si tokens(1) est bis/ter/qua → l'ajouter au numéro
            If tokens.Length > 2 AndAlso {"BIS", "TER", "QUATER", "B", "T"}.
                Contains(tokens(1).ToUpper()) Then
                numVoie &= " " & tokens(1)
                startSearch = 2
            End If
            d.NumVoie = numVoie
            ' Chercher le type de voie dans les tokens suivants
            Dim typeFound As Boolean = False
            For k As Integer = startSearch To tokens.Length - 1
                Dim tok As String = tokens(k).ToUpper()
                If typeMap.ContainsKey(tok) Then
                    d.TypeVoie = typeMap(tok)
                    Dim nomParts As New List(Of String)
                    For j As Integer = k + 1 To tokens.Length - 1
                        nomParts.Add(tokens(j))
                    Next
                    d.NomVoie = ToTitleCase(String.Join(" ", nomParts))
                    typeFound = True
                    Exit For
                End If
            Next
            ' Fallback : tokens(startSearch) = type, reste = nom
            If Not typeFound Then
                d.TypeVoie = ToTitleCase(tokens(startSearch))
                Dim nomParts As New List(Of String)
                For j As Integer = startSearch + 1 To tokens.Length - 1
                    nomParts.Add(tokens(j))
                Next
                d.NomVoie = ToTitleCase(String.Join(" ", nomParts))
            End If
        Else
            d.NomVoie = ToTitleCase(avant)
        End If
    End Sub

    Private Function ToTitleCase(s As String) As String
        If String.IsNullOrEmpty(s) Then Return s
        Return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower())
    End Function

    ''' <summary>Extrait le texte entre deux marqueurs HTML.</summary>
    Private Function HtmlVal(html As String, startMark As String, endMark As String) As String
        Dim i As Integer = html.IndexOf(startMark, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return ""
        i += startMark.Length
        Dim j As Integer = html.IndexOf(endMark, i, StringComparison.OrdinalIgnoreCase)
        If j < 0 Then Return ""
        Return System.Net.WebUtility.HtmlDecode(html.Substring(i, j - i).Trim())
    End Function

    ''' <summary>Extrait la valeur d'un attribut HTML.</summary>
    Private Function HtmlAttr(html As String, startMark As String, endMark As String) As String
        Dim i As Integer = html.IndexOf(startMark, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return ""
        i += startMark.Length
        Dim j As Integer = html.IndexOf(endMark, i)
        If j < 0 Then Return ""
        Return html.Substring(i, j - i).Trim()
    End Function

    ''' <summary>Cherche une clé JSON dans tout le HTML, gère guillemets doubles ET simples.</summary>
    Private Function JsonValHtml(html As String, key As String) As String
        For Each q As Char In {Chr(34), "'"c}
            Dim pattern As String = q & key & q
            Dim idx As Integer = html.IndexOf(pattern)
            If idx < 0 Then Continue For
            Dim colon As Integer = html.IndexOf(":", idx + pattern.Length)
            If colon < 0 Then Continue For
            Dim valStart As Integer = colon + 1
            Do While valStart < html.Length AndAlso (html(valStart) = " "c OrElse html(valStart) = vbTab)
                valStart += 1
            Loop
            If valStart >= html.Length Then Continue For
            Dim qv As Char = html(valStart)
            If qv = Chr(34) OrElse qv = "'"c Then
                valStart += 1
                Dim valEnd As Integer = html.IndexOf(qv, valStart)
                If valEnd < 0 Then Continue For
                Return html.Substring(valStart, valEnd - valStart)
            End If
        Next
        Return ""
    End Function

    Private Function JsonVal(json As String, key As String) As String
        ' Essayer guillemet double d'abord, puis apostrophe simple
        For Each q As Char In {Chr(34), "'"c}
            Dim pattern As String = q & key & q
            Dim idx As Integer = json.IndexOf(pattern)
            If idx < 0 Then Continue For
            Dim colon As Integer = json.IndexOf(":", idx + pattern.Length)
            If colon < 0 Then Continue For
            Dim valStart As Integer = colon + 1
            Do While valStart < json.Length AndAlso json(valStart) = " "c
                valStart += 1
            Loop
            If valStart >= json.Length Then Continue For
            Dim quoteChar As Char = json(valStart)
            If quoteChar = Chr(34) OrElse quoteChar = "'"c Then
                valStart += 1
                Dim valEnd As Integer = json.IndexOf(quoteChar, valStart)
                If valEnd < 0 Then Continue For
                Return json.Substring(valStart, valEnd - valStart)
            ElseIf json(valStart) = "n"c Then
                Return ""
            Else
                Dim valEnd As Integer = valStart
                Do While valEnd < json.Length AndAlso json(valEnd) <> ","c AndAlso
                         json(valEnd) <> "}"c AndAlso json(valEnd) <> "]"c
                    valEnd += 1
                Loop
                Return json.Substring(valStart, valEnd - valStart).Trim()
            End If
        Next
        Return ""
    End Function

    Private Function ExtractJsonBlock(json As String, fromIdx As Integer) As String
        Dim start As Integer = json.IndexOf("{", fromIdx)
        If start < 0 Then Return ""
        Dim depth As Integer = 0
        For i As Integer = start To json.Length - 1
            If json(i) = "{"c Then depth += 1
            If json(i) = "}"c Then
                depth -= 1
                If depth = 0 Then Return json.Substring(start, i - start + 1)
            End If
        Next
        Return ""
    End Function

End Class
