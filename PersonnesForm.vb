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
        Path.Combine(Application.StartupPath, "Data", "BDD_IQS_Person.xlsx")


    Private Const SHEET_PHY As String = "PERSONNEPHYSIQUE"
    Private Const SHEET_MOR As String = "PERSONNEMORALE"

    ' Colonnes PERSONNEPHYSIQUE
    Public Shared ReadOnly ColsPhy As String() = {
        "Pseudonyme", "Nom", "Prenom", "Genre", "Editeur", "Role",
        "COAD", "IPI", "IPI 2",
        "Num de voie", "Type de voie", "Nom de voie", "CP", "Ville",
        "Mail", "Tel", "Date de naissance", "Lieu de naissance", "N Secu"
    }

    ' Colonnes PERSONNEMORALE
    Public Shared ReadOnly ColsMor As String() = {
        "Designation", "COAD", "IPI",
        "Forme Juridique", "Capital", "RCS", "Siren",
        "Num de voie", "Type de voie", "Nom de voie", "CP", "Ville",
        "Prenom representant", "Nom representant", "Fonction representant",
        "Mail", "Tel"
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
            {"Pseudonyme", 120}, {"Nom", 110}, {"Prenom", 100}, {"Genre", 50},
            {"Editeur", 190}, {"Role", 45}, {"COAD", 80}, {"IPI", 100}, {"IPI 2", 80},
            {"Designation", 210}, {"Forme Juridique", 100}, {"Capital", 70},
            {"RCS", 80}, {"Siren", 95},
            {"Num de voie", 60}, {"Type de voie", 80}, {"Nom de voie", 150},
            {"CP", 60}, {"Ville", 110},
            {"Prenom representant", 120}, {"Nom representant", 120},
            {"Fonction representant", 130},
            {"Mail", 180}, {"Tel", 100},
            {"Date de naissance", 100}, {"Lieu de naissance", 100}, {"N Secu", 110}
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
    Private Sub BtnAddPhy_Click(sender As Object, e As EventArgs)
        Using f As New FichePersonneForm(ColsPhy, Nothing, "Nouvelle Personne Physique")
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
        Using f As New FichePersonneForm(ColsMor, Nothing, "Nouvelle Personne Morale")
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
            Else
                result(i) = _textFields(i).Text.Trim()
            End If
        Next
        Return result
    End Function

    Public Sub New(cols As String(), currentValues As String(), title As String)
        _cols = cols
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
        dropdowns("Role") = New String() {"A", "C", "AR", "AD", "E"}
        dropdowns("Forme Juridique") = New String() {"SAS", "SARL", "SA", "EURL", "EI", "SNC", "Association", "Autre"}

        For i = 0 To cols.Length - 1
            Dim y As Integer = mg + i * rowH

            Dim lbl As New Label()
            lbl.Text = cols(i) & " :"
            lbl.Location = New Point(mg, y + 5)
            lbl.Size = New Size(lblW, 20)
            lbl.TextAlign = ContentAlignment.MiddleRight
            pnl.Controls.Add(lbl)

            Dim colKey As String = cols(i)
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

        AddHandler btnOK.Click, AddressOf BtnOK_Click
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
