Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Data
Imports Newtonsoft.Json.Linq
Imports OfficeOpenXml

' ════════════════════════════════════════════════════════════════════════════
' FormRecherchePublicSACEM.vb
' Répertoire Public SACEM — recherche + enrichissement détail
' ════════════════════════════════════════════════════════════════════════════
Public Class FormRecherchePublicSACEM
    Inherits Form

    ' ── Couleurs ──────────────────────────────────────────────────────────
    Private Shared ReadOnly C_HEAD As Color = Color.FromArgb(20, 60, 140)
    Private Shared ReadOnly C_BG As Color = Color.FromArgb(230, 240, 255)
    Private Shared ReadOnly C_OK As Color = Color.FromArgb(16, 124, 16)
    Private Shared ReadOnly C_STOP As Color = Color.FromArgb(196, 43, 28)
    Private Shared ReadOnly C_SAVE As Color = Color.FromArgb(0, 100, 80)
    Private Shared ReadOnly C_IPI As Color = Color.FromArgb(195, 230, 255)
    Private Shared ReadOnly C_FILT As Color = Color.FromArgb(255, 255, 210)

    ' ── Chemins ───────────────────────────────────────────────────────────
    Public Shared ReadOnly ScriptPath As String =
        Path.Combine(Application.StartupPath, "..", "..", "Scripts", "sacem_repertoire_public.py")
    Public Shared ReadOnly BddPath As String =
        Path.Combine(Application.StartupPath, "..", "..", "Data", "BDD_SACEM_Repertoire.xlsx")

    ' ── Colonnes ──────────────────────────────────────────────────────────
    Private Shared ReadOnly COL_HEADERS As New Dictionary(Of String, String) From {
        {"Sel", ""}, {"Titre", "Titre"}, {"ISWC", "ISWC"}, {"Type", "Type"}, {"Duree", "Durée"},
        {"A", "Auteur(s)"}, {"C", "Compositeur(s)"}, {"AR", "Arrangeur(s)"},
        {"AD", "Adaptateur(s)"}, {"E", "Éditeur(s)"}, {"SE", "Sous-éditeur(s)"},
        {"Interp", "Interprète(s)"}, {"RE", "Réalisateur(s)"}
    }

    ' ── État ──────────────────────────────────────────────────────────────
    Private _process As Process
    Private _oeuvres As New List(Of JObject)()
    Private _tokensConnus As New HashSet(Of String)()          ' clés ISWC ou titre — déduplication inter-requêtes
    Private _dt As DataTable
    Private _bs As BindingSource
    Private _total As Integer = 0
    Private _maxPage As Integer = 0
    Private _running As Boolean = False
    Private _cancelRequested As Boolean = False
    Private _sortCol As String = ""
    Private _sortAsc As Boolean = True
    Private _enrichTotal As Integer = 0
    Private _enrichDone As Integer = 0
    Private _detailsBuffer As New Dictionary(Of Integer, JObject)()
    Private _modeDetails As Boolean = False
    Private _ignoreNextDone As Boolean = False   ' True = ignorer le prochain "done" (mode détails)
    Private _enrichOffset As Integer = 0  ' nb d'oeuvres déjà dans _dt avant la requête en cours

    ' Stats par requête : (nom, nouvelles, doublons)
    Private _statsParRequete As New List(Of Tuple(Of String, Integer, Integer, Integer))()  ' nom, nouv, dbl, totalSACEM
    Private _currentReqNom As String = ""
    Private _currentReqTotal As Integer = 0   ' total annoncé par SACEM pour la requête en cours
    Private _currentReqNouv As Integer = 0
    Private _currentReqDbl As Integer = 0

    ' ── Contrôles ─────────────────────────────────────────────────────────
    Private pnlHeader As Panel
    Private txtTitre As TextBox
    Private txtCreateur As TextBox   ' champ créateur (supporte + comme séparateur)
    Private _queryQueue As New Queue(Of String)()
    Private _planRequetes As New List(Of String)()   ' toutes les queries dans l'ordre, connu dès BtnRechercher_Click
    Private _statTotaux As New Dictionary(Of String, Integer)()  ' query → total SACEM (pré-stat)
    Private _pendingFiltre As String = ""
    Private _pendingTitre As String = ""
    Private _lastQuery As String = ""   ' query de la dernière recherche (pour --details-seulement)
    Private _lastFiltre As String = "parties"
    Private _ipiRecherche As String = ""   ' IPI filter courant
    Private btnRechercher As Button
    Private btnAnnuler As Button
    Private btnSauvegarder As Button
    Private btnFiltrer As Button
    Private btnExtraire As Button
    Private btnDetails As Button   ' actif quand tous les détails sont récupérés
    Private pnlProgress As Panel
    Private lblProgress As Label
    Private pbProgress As ProgressBar
    Private pnlRecherche As Panel        ' barre de recherche globale (activée après enrichissement)
    Private txtRecherche As TextBox      ' recherche multi-termes toutes colonnes
    Private lblStats As Label
    Private lblRechInfo As Label        ' "n résultats"
    Private pnlFiltres As Panel
    Private _filterBoxes As New Dictionary(Of String, TextBox)()
    Private dgvOeuvres As DataGridView
    Private lblInfo As Label
    Private txtLog As TextBox

    ' ══════════════════════════════════════════════════════════════════════
    Public Sub New()
        InitializeComponent()
        BuildDataTable()
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' INIT UI  (ordre Add = inverse ordre visuel pour Dock)
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub InitializeComponent()
        Me.Text = "Répertoire Public SACEM"
        Me.Size = New Size(1420, 820)
        Me.MinimumSize = New Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = C_BG

        ' ── Bottom ────────────────────────────────────────────────────
        lblInfo = New Label()
        lblInfo.Dock = DockStyle.Bottom
        lblInfo.Height = 20
        lblInfo.Font = New Font("Segoe UI", 8.0F)
        lblInfo.ForeColor = Color.FromArgb(60, 80, 120)
        lblInfo.BackColor = Color.FromArgb(215, 228, 250)
        lblInfo.Padding = New Padding(8, 2, 0, 0)
        lblInfo.Text = "Prêt."
        Me.Controls.Add(lblInfo)

        txtLog = New TextBox()
        txtLog.Dock = DockStyle.Bottom
        txtLog.Height = 70
        txtLog.Multiline = True
        txtLog.ReadOnly = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.Font = New Font("Consolas", 7.5F)
        txtLog.ForeColor = Color.FromArgb(20, 60, 20)
        txtLog.BackColor = Color.FromArgb(240, 248, 240)
        txtLog.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(txtLog)

        ' ── Fill : DGV ────────────────────────────────────────────────
        dgvOeuvres = New DataGridView()
        dgvOeuvres.Dock = DockStyle.Fill
        dgvOeuvres.AllowUserToAddRows = False
        dgvOeuvres.AllowUserToDeleteRows = False
        dgvOeuvres.AllowUserToOrderColumns = True
        dgvOeuvres.AllowDrop = True
        dgvOeuvres.ReadOnly = False
        dgvOeuvres.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgvOeuvres.MultiSelect = True
        dgvOeuvres.ScrollBars = ScrollBars.Both
        dgvOeuvres.BackgroundColor = Color.FromArgb(235, 244, 255)
        dgvOeuvres.GridColor = Color.FromArgb(190, 210, 240)
        dgvOeuvres.BorderStyle = BorderStyle.None
        dgvOeuvres.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgvOeuvres.ColumnHeadersHeight = 28
        dgvOeuvres.EnableHeadersVisualStyles = False
        dgvOeuvres.ColumnHeadersDefaultCellStyle.BackColor = C_HEAD
        dgvOeuvres.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvOeuvres.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        dgvOeuvres.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvOeuvres.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(222, 234, 252)
        dgvOeuvres.AutoGenerateColumns = False
        dgvOeuvres.EditMode = DataGridViewEditMode.EditOnF2
        dgvOeuvres.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        AddHandler dgvOeuvres.CellClick, AddressOf Dgv_CellClick
        AddHandler dgvOeuvres.CellContentClick, AddressOf Dgv_CellContentClick
        AddHandler dgvOeuvres.KeyDown, AddressOf Dgv_KeyDown
        AddHandler dgvOeuvres.CellDoubleClick, AddressOf Dgv_CellDoubleClick
        AddHandler dgvOeuvres.CellMouseEnter, AddressOf Dgv_CellMouseEnter
        AddHandler dgvOeuvres.ColumnHeaderMouseClick, AddressOf Dgv_HeaderClick
        AddHandler dgvOeuvres.SortCompare, AddressOf Dgv_SortCompare
        AddHandler dgvOeuvres.Scroll, AddressOf Dgv_Scroll
        AddHandler dgvOeuvres.ColumnWidthChanged, AddressOf Dgv_ColChanged
        AddHandler dgvOeuvres.ColumnDisplayIndexChanged, AddressOf Dgv_ColChanged
        AddHandler dgvOeuvres.MouseDown, AddressOf Dgv_RowDrag_MouseDown
        AddHandler dgvOeuvres.MouseMove, AddressOf Dgv_RowDrag_MouseMove
        AddHandler dgvOeuvres.DragOver, AddressOf Dgv_RowDrag_DragOver
        AddHandler dgvOeuvres.DragDrop, AddressOf Dgv_RowDrag_DragDrop
        ' ── Top : ligne filtres (dans pnlCenter)
        pnlFiltres = New Panel()
        pnlFiltres.Dock = DockStyle.Top
        pnlFiltres.Height = 26
        pnlFiltres.BackColor = C_FILT

        ' ── Top : barre recherche globale ─────────────────────────────────
        pnlRecherche = New Panel()
        pnlRecherche.Dock = DockStyle.Top
        pnlRecherche.Height = 30
        pnlRecherche.BackColor = Color.FromArgb(220, 235, 255)
        ' pnlRecherche ajouté dans pnlCenter plus bas

        Dim lbR As New Label()
        lbR.Text = "🔎"
        lbR.Font = New Font("Segoe UI", 9.0F)
        lbR.ForeColor = Color.FromArgb(60, 80, 140)
        lbR.Location = New Point(6, 6) : lbR.AutoSize = True
        pnlRecherche.Controls.Add(lbR)

        txtRecherche = New TextBox()
        txtRecherche.Location = New Point(26, 4)
        txtRecherche.Size = New Size(420, 22)
        txtRecherche.Font = New Font("Segoe UI", 9.0F)
        txtRecherche.ForeColor = Color.Gray
        txtRecherche.Text = "IPI ou nom… (!A = exclure, A + B = ET, A * B = OU)"
        AddHandler txtRecherche.GotFocus, AddressOf TxtRecherche_GotFocus
        AddHandler txtRecherche.LostFocus, AddressOf TxtRecherche_LostFocus
        AddHandler txtRecherche.KeyDown, AddressOf TxtRecherche_KeyDown
        pnlRecherche.Controls.Add(txtRecherche)

        btnFiltrer = MkBtn("Filtrer", 452, 2, 68, Color.FromArgb(40, 80, 160))
        AddHandler btnFiltrer.Click, AddressOf BtnFiltrer_Click
        pnlRecherche.Controls.Add(btnFiltrer)

        btnExtraire = MkBtn("📋 Extraire rôles", 526, 2, 130, Color.FromArgb(80, 40, 140))
        AddHandler btnExtraire.Click, AddressOf BtnExtraire_Click
        pnlRecherche.Controls.Add(btnExtraire)

        ' Label stats permanent — mis à jour par MettreAJourLblInfo
        lblStats = New Label()
        lblStats.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        lblStats.ForeColor = Color.FromArgb(20, 60, 130)
        lblStats.BackColor = Color.FromArgb(200, 218, 250)
        lblStats.Padding = New Padding(6, 0, 6, 0)
        lblStats.Location = New Point(670, 4)
        lblStats.Height = 22
        lblStats.AutoSize = True
        lblStats.Anchor = AnchorStyles.Left Or AnchorStyles.Top
        lblStats.Text = ""
        pnlRecherche.Controls.Add(lblStats)

        lblRechInfo = New Label()
        lblRechInfo.Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        lblRechInfo.ForeColor = Color.FromArgb(0, 90, 40)
        lblRechInfo.Location = New Point(148, 8)
        lblRechInfo.AutoSize = True
        lblRechInfo.Visible = False
        pnlRecherche.Controls.Add(lblRechInfo)

        ' ── Top : barre progression ────────────────────────────────────
        pnlProgress = New Panel()
        pnlProgress.Dock = DockStyle.Top
        pnlProgress.Height = 24
        pnlProgress.BackColor = Color.FromArgb(210, 225, 248)
        pnlProgress.Visible = False
        ' pnlProgress ajouté dans pnlCenter plus bas

        pbProgress = New ProgressBar()
        pbProgress.Location = New Point(6, 3)
        pbProgress.Size = New Size(600, 18)
        pbProgress.Style = ProgressBarStyle.Marquee
        pnlProgress.Controls.Add(pbProgress)

        lblProgress = New Label()
        lblProgress.Location = New Point(614, 4)
        lblProgress.Size = New Size(700, 16)
        lblProgress.Font = New Font("Segoe UI", 8.0F)
        lblProgress.ForeColor = Color.FromArgb(30, 60, 140)
        pnlProgress.Controls.Add(lblProgress)

        ' ── Top : bandeau (ajouté en dernier = tout en haut) ──────────
        pnlHeader = New Panel()
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.Height = 62
        pnlHeader.BackColor = C_HEAD
        ' pnlHeader ajouté après pnlCenter

        ' Ligne 1 : label source
        AddLbl(pnlHeader, "RÉPERTOIRE PUBLIC SACEM  ·  repertoire.sacem.fr",
               New Font("Segoe UI", 8.0F, FontStyle.Bold), Color.FromArgb(160, 200, 255), 10, 4)

        ' Ligne 2 : Titre + Créateur + boutons
        AddLbl(pnlHeader, "Titre :", New Font("Segoe UI", 8.5F), Color.White, 10, 28)
        txtTitre = New TextBox()
        txtTitre.Location = New Point(55, 26)
        txtTitre.Size = New Size(260, 23)
        txtTitre.Font = New Font("Segoe UI", 9.0F)
        AddHandler txtTitre.KeyDown, AddressOf TxtQuery_KeyDown
        pnlHeader.Controls.Add(txtTitre)

        AddLbl(pnlHeader, "Créateur (+ = 2 requêtes) :", New Font("Segoe UI", 8.5F), Color.White, 325, 28)
        txtCreateur = New TextBox()
        txtCreateur.Location = New Point(510, 26)
        txtCreateur.Size = New Size(140, 23)
        txtCreateur.Font = New Font("Segoe UI", 9.0F)
        AddHandler txtCreateur.KeyDown, AddressOf TxtQuery_KeyDown
        pnlHeader.Controls.Add(txtCreateur)

        btnRechercher = MkBtn("🔍 Rechercher", 662, 26, 130, C_OK)
        AddHandler btnRechercher.Click, AddressOf BtnRechercher_Click
        pnlHeader.Controls.Add(btnRechercher)

        btnAnnuler = MkBtn("■ Stop", 798, 26, 68, C_STOP)
        btnAnnuler.Enabled = False
        AddHandler btnAnnuler.Click, AddressOf BtnAnnuler_Click
        pnlHeader.Controls.Add(btnAnnuler)

        btnSauvegarder = MkBtn("💾 Sauvegarder", 876, 26, 122, C_SAVE)
        AddHandler btnSauvegarder.Click, AddressOf BtnSauvegarder_Click
        pnlHeader.Controls.Add(btnSauvegarder)

        btnDetails = MkBtn("Détails", 1006, 26, 110, Color.FromArgb(80, 120, 60))
        btnDetails.Enabled = False
        AddHandler btnDetails.Click, AddressOf BtnDetails_Click
        pnlHeader.Controls.Add(btnDetails)

        ' pnlCenter = conteneur Fill : DGV (Fill) + tous les Top
        Dim pnlCenter As New Panel()
        pnlCenter.Dock = DockStyle.Fill
        pnlCenter.Controls.Add(dgvOeuvres)   ' Fill en premier
        pnlCenter.Controls.Add(pnlFiltres)   ' Top — filtre par colonne
        pnlCenter.Controls.Add(pnlRecherche) ' Top — recherche globale
        pnlCenter.Controls.Add(pnlProgress)  ' Top — barre de progression
        Me.Controls.Add(pnlCenter)
        Me.Controls.Add(pnlHeader)

        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler Me.FormClosing, AddressOf Form_Closing
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' DATATABLE + COLONNES DGV
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub BuildDataTable()
        _dt = New DataTable()
        _dt.Columns.Add("Sel", GetType(Boolean))
        For Each nom In {"Titre", "ISWC", "Type", "Duree", "Interp", "A", "RE", "C", "E", "SE", "AR", "AD", "SousTitres", "Token"}
            _dt.Columns.Add(nom, GetType(String))
        Next
        _dt.Columns.Add("IpiMatch", GetType(Boolean))

        _bs = New BindingSource()
        _bs.DataSource = _dt
        AddHandler _dt.RowChanged, AddressOf Dt_RowChanged

        dgvOeuvres.Columns.Clear()

        ' ☐ Sel
        Dim cSel As New DataGridViewCheckBoxColumn()
        cSel.Name = "Sel" : cSel.HeaderText = "☐"
        cSel.DataPropertyName = "Sel"
        cSel.Width = 28 : cSel.ReadOnly = False
        cSel.SortMode = DataGridViewColumnSortMode.NotSortable
        cSel.ToolTipText = "Clic entête = tout cocher / décocher"
        dgvOeuvres.Columns.Add(cSel)

        ' Titre
        Dim cTitre As New DataGridViewTextBoxColumn()
        cTitre.Name = "Titre" : cTitre.HeaderText = "Titre"
        cTitre.DataPropertyName = "Titre"
        cTitre.MinimumWidth = 120 : cTitre.ReadOnly = True
        cTitre.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        cTitre.SortMode = DataGridViewColumnSortMode.Programmatic
        cTitre.DefaultCellStyle.ForeColor = Color.FromArgb(0, 80, 200)
        cTitre.DefaultCellStyle.Font = New Font("Segoe UI", 8.5F, FontStyle.Underline)
        cTitre.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        cTitre.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        dgvOeuvres.Columns.Add(cTitre)

        ' Colonnes visibles dans l'ordre voulu
        ' Titre déjà ajouté — ajouter ISWC Type Duree Interp A RE C E SE AR AD
        For Each kv In {("ISWC", "ISWC", 80), ("Type", "Type", 80), ("Duree", "Durée", 60),
                        ("Interp", "Interprète(s)", 160), ("A", "Auteur(s)", 160), ("RE", "Réalisateur(s)", 160),
                        ("C", "Compositeur(s)", 160), ("E", "Éditeur(s)", 200), ("SE", "Sous-éditeur(s)", 200),
                        ("AR", "Arrangeur(s)", 160), ("AD", "Adaptateur(s)", 160),
                        ("SousTitres", "Sous-titre(s)", 200)}
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = kv.Item1 : col.HeaderText = kv.Item2
            col.DataPropertyName = kv.Item1
            col.MinimumWidth = 40 : col.ReadOnly = True
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            col.SortMode = DataGridViewColumnSortMode.Programmatic
            col.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            col.DefaultCellStyle.Font = New Font("Segoe UI", 8.0F)
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvOeuvres.Columns.Add(col)
        Next

        ' Cachées
        For Each nom In {"Token", "IpiMatch"}
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = nom : col.DataPropertyName = nom : col.Visible = False
            dgvOeuvres.Columns.Add(col)
        Next

        dgvOeuvres.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        dgvOeuvres.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgvOeuvres.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvOeuvres.DataSource = _bs
    End Sub

    ' Coloration IPI match + réactivation couleur colonnes détail après enrichissement
    Private Sub Dt_RowChanged(sender As Object, e As DataRowChangeEventArgs)
        Dim rowIdx As Integer = _dt.Rows.IndexOf(e.Row)
        If rowIdx < 0 OrElse rowIdx >= dgvOeuvres.RowCount Then Return
        If CBool(If(e.Row("IpiMatch"), False)) Then
            dgvOeuvres.Rows(rowIdx).DefaultCellStyle.BackColor = C_IPI
            dgvOeuvres.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.FromArgb(0, 40, 100)
        End If
        If e.Action = DataRowAction.Change Then MajLblInfo()
    End Sub

    ' Appelée après enrichissement phase 2
    Private Sub ActiverCouleursDetail()
        lblRechInfo.Visible = False
        dgvOeuvres.Refresh()
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' LIGNE DE FILTRES PAR COLONNE
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub BuildFilterRow()
        pnlFiltres.Controls.Clear()
        _filterBoxes.Clear()
        For Each col As DataGridViewColumn In dgvOeuvres.Columns
            If Not col.Visible OrElse col.Name = "Sel" Then Continue For
            Dim tb As New TextBox()
            tb.Font = New Font("Segoe UI", 7.5F)
            tb.BackColor = C_FILT
            tb.BorderStyle = BorderStyle.None
            tb.ForeColor = Color.FromArgb(60, 60, 120)
            tb.Tag = col.Name
            AddHandler tb.TextChanged, AddressOf AppliquerFiltres
            pnlFiltres.Controls.Add(tb)
            _filterBoxes(col.Name) = tb
        Next
        PositionFilters()
    End Sub

    Private Sub PositionFilters()
        If _filterBoxes.Count = 0 OrElse Not dgvOeuvres.IsHandleCreated Then Return
        Dim scroll As Integer = dgvOeuvres.HorizontalScrollingOffset
        Dim x As Integer = dgvOeuvres.RowHeadersWidth - scroll
        For Each col As DataGridViewColumn In dgvOeuvres.Columns
            If Not col.Visible Then Continue For
            If col.Name = "Sel" Then x += col.Width : Continue For
            If _filterBoxes.ContainsKey(col.Name) Then
                Dim tb As TextBox = _filterBoxes(col.Name)
                tb.Location = New Point(x + 2, 4)
                tb.Height = 18
                tb.Width = Math.Max(col.Width - 4, 10)
            End If
            x += col.Width
        Next
    End Sub

    ' Appelé par FilterBox (par colonne) ET par txtRecherche (global)
    Private Sub AppliquerFiltres(sender As Object, e As EventArgs)
        Dim colParts As New List(Of String)()

        ' ── Filtres par colonne ───────────────────────────────────────
        For Each kvp In _filterBoxes
            Dim v As String = kvp.Value.Text.Trim()
            If Not String.IsNullOrEmpty(v) Then
                colParts.Add($"CONVERT([{kvp.Key}], System.String) LIKE '%{v.Replace("'", "''")}%'")
            End If
        Next

        ' ── Recherche globale : +  = ET  /  * = OU  /  !terme = exclure
        ' Priorité : + avant *
        ' Ex : "!438009074 + KOBALT * MOHOMBI"
        '      → (NE contient PAS 438009074 ET contient KOBALT) OU contient MOHOMBI
        Dim rechTexte As String = If(Not _rechWatermark,
                                     txtRecherche.Text.Trim(), "")

        If Not String.IsNullOrEmpty(rechTexte) Then
            Dim cols As New List(Of String)()
            For Each col As DataGridViewColumn In dgvOeuvres.Columns
                If col.Visible AndAlso Not col.Name = "Sel" AndAlso
                   Not col.Name = "Token" AndAlso Not col.Name = "IpiMatch" Then
                    cols.Add(col.Name)
                End If
            Next

            ' Séparer en groupes OU (séparateur *)
            Dim groupesOU As String() = rechTexte.Split(
                New Char() {"*"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim ouParts As New List(Of String)()

            For Each groupe In groupesOU
                ' Séparer en termes ET (séparateur +)
                Dim termes As String() = groupe.Split(
                    New Char() {"+"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim etParts As New List(Of String)()

                For Each terme In termes
                    Dim t As String = terme.Trim()
                    If String.IsNullOrEmpty(t) Then Continue For

                    Dim exclure As Boolean = t.StartsWith("!")
                    If exclure Then t = t.Substring(1).Trim()
                    If String.IsNullOrEmpty(t) Then Continue For
                    t = t.Replace("'", "''")

                    Dim colExprs As New List(Of String)()
                    For Each nom In cols
                        colExprs.Add($"CONVERT([{nom}], System.String) LIKE '%{t}%'")
                    Next
                    If colExprs.Count = 0 Then Continue For

                    If exclure Then
                        ' NOT (col1 LIKE OR col2 LIKE ...) = aucune colonne ne contient le terme
                        etParts.Add("NOT (" & String.Join(" OR ", colExprs) & ")")
                    Else
                        etParts.Add("(" & String.Join(" OR ", colExprs) & ")")
                    End If
                Next

                If etParts.Count > 0 Then
                    ouParts.Add("(" & String.Join(" AND ", etParts) & ")")
                End If
            Next

            If ouParts.Count > 0 Then
                colParts.Add("(" & String.Join(" OR ", ouParts) & ")")
            End If
        End If

        Try
            _dt.DefaultView.RowFilter = String.Join(" AND ", colParts)
        Catch
            _dt.DefaultView.RowFilter = ""
        End Try

        Dim cnt As Integer = _dt.DefaultView.Count
        Dim tot As Integer = _dt.Rows.Count
        lblInfo.Text = $"{cnt}/{tot} oeuvre(s)"
        MajLblInfo()

        If Not String.IsNullOrEmpty(rechTexte) Then
            lblRechInfo.Text = $"{cnt} résultat(s)"
            lblRechInfo.Visible = True
        Else
            lblRechInfo.Visible = False
        End If
    End Sub

    ' ── Watermark / handlers txtRecherche ─────────────────────────────
    Private _rechWatermark As Boolean = True

    Private Sub TxtRecherche_GotFocus(sender As Object, e As EventArgs)
        If _rechWatermark Then
            _rechWatermark = False
            txtRecherche.Text = ""
            txtRecherche.ForeColor = Color.FromArgb(20, 40, 90)
        End If
    End Sub

    Private Sub TxtRecherche_LostFocus(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtRecherche.Text) Then
            _rechWatermark = True
            txtRecherche.Text = "IPI ou nom… (!A = exclure, A + B = ET, A * B = OU)"
            txtRecherche.ForeColor = Color.Gray
            AppliquerFiltres(Nothing, EventArgs.Empty)
        End If
    End Sub

    Private Sub BtnFiltrer_Click(sender As Object, e As EventArgs)
        AppliquerFiltres(Nothing, EventArgs.Empty)
    End Sub

    Private Sub BtnDetails_Click(sender As Object, e As EventArgs)
        If _oeuvres.Count = 0 Then Return
        btnDetails.Enabled = False
        btnDetails.Text = "⏳ Chargement…"

        ' Construire la liste TOKEN:TITRE:INDEX → fichier temp (ligne de commande trop longue sinon)
        Dim items As New List(Of String)()
        For i As Integer = 0 To _oeuvres.Count - 1
            Dim oe As JObject = _oeuvres(i)
            Dim tok As String = If(CStr(oe("token")), "")
            Dim tit As String = If(CStr(oe("titre")), "").Replace(":", " ").Replace("""", " ").Replace(Chr(10), " ")
            If Not String.IsNullOrEmpty(tok) Then
                items.Add($"{tok}:{tit}:{i}")
            End If
        Next

        ' Contenu stdin : un TOKEN:TITRE:INDEX par ligne
        Dim stdinContent As String = String.Join(Environment.NewLine, items)

        Dim detailsArg As String = " --details-stdin"

        ' Filtre IPI si présent
        Dim ipiArg As String = ""
        If Not String.IsNullOrEmpty(_ipiRecherche) Then
            ipiArg = $" --ipi {_ipiRecherche}"
        End If

        _modeDetails = True
        SetRunningState(True)
        lblProgress.Text = $"Récupération des détails (0/{_oeuvres.Count})…"
        pbProgress.Maximum = _oeuvres.Count
        pbProgress.Value = 0
        _enrichDone = 0
        _enrichTotal = _oeuvres.Count

        Dim fullArgs As String = $" --query ""{_lastQuery}"" --filtre {_lastFiltre}{ipiArg}{detailsArg}"
        AddLog($"[DETAILS] {items.Count} tokens via stdin")
        LancerProcessAvecStdin(fullArgs, stdinContent)
    End Sub

    Private Sub TxtRecherche_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            AppliquerFiltres(Nothing, EventArgs.Empty)
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Escape Then
            txtRecherche.Text = ""
            TxtRecherche_LostFocus(sender, EventArgs.Empty)
            e.SuppressKeyPress = True
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' TRI
    ' ══════════════════════════════════════════════════════════════════════
    Private _selToggle As Boolean = True   ' True = prochain clic cochera tout, False = décochera

    Private Sub Dgv_HeaderClick(sender As Object, e As DataGridViewCellMouseEventArgs)
        Dim col As DataGridViewColumn = dgvOeuvres.Columns(e.ColumnIndex)

        ' Clic sur entête Sel = tout cocher / décocher (lignes visibles seulement)
        If col.Name = "Sel" Then
            For Each drv As DataRowView In _bs
                drv.Row("Sel") = _selToggle
            Next
            col.HeaderText = If(_selToggle, "☑", "☐")
            _selToggle = Not _selToggle
            dgvOeuvres.Refresh()
            Return
        End If

        If col.SortMode = DataGridViewColumnSortMode.NotSortable Then Return
        _sortAsc = If(_sortCol = col.Name, Not _sortAsc, True)
        _sortCol = col.Name
        For Each c As DataGridViewColumn In dgvOeuvres.Columns
            c.HeaderCell.SortGlyphDirection = SortOrder.None
        Next
        col.HeaderCell.SortGlyphDirection = If(_sortAsc, SortOrder.Ascending, SortOrder.Descending)
        dgvOeuvres.Sort(col, If(_sortAsc,
            System.ComponentModel.ListSortDirection.Ascending,
            System.ComponentModel.ListSortDirection.Descending))
    End Sub

    Private Sub Dgv_SortCompare(sender As Object, e As DataGridViewSortCompareEventArgs)
        Dim v1 As String = PremiereLigne(TryCast(e.CellValue1, String))
        Dim v2 As String = PremiereLigne(TryCast(e.CellValue2, String))
        e.SortResult = String.Compare(v1, v2, StringComparison.CurrentCultureIgnoreCase)
        e.Handled = True
    End Sub

    Private Shared Function PremiereLigne(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim idx As Integer = s.IndexOfAny(New Char() {Chr(10), Chr(13)})
        Return If(idx >= 0, s.Substring(0, idx).Trim(), s.Trim())
    End Function

    ' ══════════════════════════════════════════════════════════════════════
    ' LOAD / RESIZE / CLOSING
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub Form_Load(sender As Object, e As EventArgs)
        If Not File.Exists(Path.GetFullPath(ScriptPath)) Then
            lblInfo.Text = "Script introuvable : " & Path.GetFullPath(ScriptPath)
        End If
        BuildFilterRow()
        txtTitre.Focus()
    End Sub

    Private Sub Form_Resize(sender As Object, e As EventArgs)
        PositionFilters()
    End Sub

    Private Sub Form_Closing(sender As Object, e As FormClosingEventArgs)
        _cancelRequested = True
        KillProcess()
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' RECHERCHE
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub BtnRechercher_Click(sender As Object, e As EventArgs)
        Dim qTitre As String = txtTitre.Text.Trim()
        Dim qCreateur As String = txtCreateur.Text.Trim()

        If String.IsNullOrEmpty(qTitre) AndAlso String.IsNullOrEmpty(qCreateur) Then
            MessageBox.Show("Saisissez un titre ou un créateur." & vbLf &
                            "Utilisez + pour plusieurs créateurs : ""Elie Yaffa + Booba""",
                            "Recherche", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        End If
        If _running Then Return

        _oeuvres.Clear() : _dt.Rows.Clear() : _tokensConnus.Clear()
        _total = 0 : _maxPage = 0 : _cancelRequested = False
        _enrichTotal = 0 : _enrichDone = 0 : _enrichOffset = 0
        _detailsBuffer.Clear()
        _modeDetails = False
        _ignoreNextDone = False
        btnDetails.Enabled = False
        btnDetails.Text = "Détails"
        lblStats.Text = ""
        _statsParRequete.Clear() : _planRequetes.Clear() : _currentReqNom = "" : _currentReqNouv = 0 : _currentReqDbl = 0 : _currentReqTotal = 0
        _dt.DefaultView.RowFilter = ""
        For Each tb As TextBox In _filterBoxes.Values : tb.Text = "" : Next

        lblInfo.Text = "Recherche en cours…" : txtLog.Clear()
        _queryQueue.Clear()

        If Not String.IsNullOrEmpty(qCreateur) Then
            ' Séparer créateur sur + → autant de requêtes que de segments
            Dim segments = qCreateur.Split("+"c).Select(Function(s) s.Trim()).
                                     Where(Function(s) Not String.IsNullOrEmpty(s)).ToList()
            If Not String.IsNullOrEmpty(qTitre) Then
                ' Titre+Créateur → filtre "titles,parties", query="TITRE,CREATEUR"
                For Each seg In segments
                    _queryQueue.Enqueue(qTitre & "," & seg)
                    _planRequetes.Add(seg)
                Next
                AddLog($"[QUEUE] {_queryQueue.Count} requête(s) titre+créateur")
                _pendingFiltre = "titles,parties"
                _pendingTitre = qTitre
                StatRecherche(Sub()
                                  SetRunningState(True)
                                  LancerProchainerequete("titles,parties", qTitre)
                              End Sub)
            Else
                ' Créateur seul → filtre "parties"
                For Each seg In segments
                    _queryQueue.Enqueue(seg)
                    _planRequetes.Add(seg)
                Next
                AddLog($"[QUEUE] {_queryQueue.Count} requête(s) créateur")
                _pendingFiltre = "parties"
                _pendingTitre = ""
                StatRecherche(Sub()
                                  SetRunningState(True)
                                  LancerProchainerequete("parties", "")
                              End Sub)
            End If
        Else
            ' Titre seul → filtre "title"
            AddLog($"[START] titre=""{qTitre}""")
            _pendingFiltre = "title"
            _pendingTitre = ""
            _lastQuery = qTitre
            _lastFiltre = "title"
            _planRequetes.Add(qTitre)
            StatRecherche(Sub()
                              SetRunningState(True)
                              LancerProcess($" --query ""{qTitre}"" --filtre title --sans-details")
                          End Sub)
        End If
    End Sub

    Private Sub LancerProchainerequete(filtre As String, qTitre As String)
        If _cancelRequested OrElse _queryQueue.Count = 0 Then
            OnProcessTermine() : Return
        End If
        Dim query As String = _queryQueue.Dequeue()
        _enrichOffset = _oeuvres.Count   ' mémoriser le nb d'oeuvres déjà présentes
        _currentReqNom = query
        _currentReqNouv = 0
        _currentReqDbl = 0
        _currentReqTotal = 0
        AddLog($"[START] query=""{query}""  filtre={filtre}  offset={_enrichOffset}  ({_queryQueue.Count} restante(s))")
        _pendingFiltre = filtre
        _pendingTitre = qTitre
        _lastQuery = query
        _lastFiltre = filtre

        LancerProcess($" --query ""{query}"" --filtre {filtre} --sans-details")
    End Sub

    Private Sub TxtQuery_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            BtnRechercher_Click(sender, EventArgs.Empty)
            e.SuppressKeyPress = True
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' PROCESS PYTHON
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub LancerProcessAvecStdin(args As String, stdinContent As String)
        Dim psi As New ProcessStartInfo()
        psi.FileName = "python"
        psi.Arguments = $"""{Path.GetFullPath(ScriptPath)}"" {args}"
        psi.UseShellExecute = False
        psi.RedirectStandardOutput = True
        psi.RedirectStandardError = True
        psi.RedirectStandardInput = True
        psi.StandardOutputEncoding = Encoding.UTF8
        psi.StandardErrorEncoding = Encoding.UTF8
        psi.CreateNoWindow = True

        _process = New Process()
        _process.StartInfo = psi
        _process.EnableRaisingEvents = True
        AddHandler _process.OutputDataReceived, AddressOf Process_Out
        AddHandler _process.ErrorDataReceived, AddressOf Process_Err
        AddHandler _process.Exited, AddressOf Process_Exit

        Try
            _process.Start()
            ' Écrire les tokens dans stdin puis fermer
            _process.StandardInput.Write(stdinContent)
            _process.StandardInput.Close()
            _process.BeginOutputReadLine()
            _process.BeginErrorReadLine()
        Catch ex As Exception
            SetRunningState(False)
            lblInfo.Text = "Erreur Python : " & ex.Message
        End Try
    End Sub

    Private Sub LancerProcess(args As String)
        Dim psi As New ProcessStartInfo()
        psi.FileName = "python"
        psi.Arguments = $"""{Path.GetFullPath(ScriptPath)}"" {args}"
        psi.UseShellExecute = False
        psi.RedirectStandardOutput = True
        psi.RedirectStandardError = True
        psi.StandardOutputEncoding = Encoding.UTF8
        psi.StandardErrorEncoding = Encoding.UTF8
        psi.CreateNoWindow = True

        _process = New Process()
        _process.StartInfo = psi
        _process.EnableRaisingEvents = True
        AddHandler _process.OutputDataReceived, AddressOf Process_Out
        AddHandler _process.ErrorDataReceived, AddressOf Process_Err
        AddHandler _process.Exited, AddressOf Process_Exit

        Try
            _process.Start()
            _process.BeginOutputReadLine()
            _process.BeginErrorReadLine()
        Catch ex As Exception
            SetRunningState(False)
            lblInfo.Text = "Erreur Python : " & ex.Message
        End Try
    End Sub

    Private Sub Process_Out(sender As Object, e As DataReceivedEventArgs)
        If e.Data Is Nothing OrElse _cancelRequested Then Return
        Dim line As String = e.Data.Trim()
        If String.IsNullOrEmpty(line) Then Return
        Dim jo As JObject
        Try : jo = JObject.Parse(line) : Catch : Return : End Try
        If Me.IsHandleCreated Then Me.BeginInvoke(Sub() TraiterLigne(jo))
    End Sub

    Private Sub Process_Err(sender As Object, e As DataReceivedEventArgs)
        If e.Data Is Nothing OrElse String.IsNullOrEmpty(e.Data.Trim()) Then Return
        If Me.IsHandleCreated Then Me.BeginInvoke(Sub() lblProgress.Text = e.Data.Trim())
    End Sub

    Private Sub Process_Exit(sender As Object, e As EventArgs)
        If Me.IsHandleCreated Then Me.BeginInvoke(Sub() OnProcessTermine())
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' TRAITEMENT JSON
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub TraiterLigne(jo As JObject)
        Select Case jo("type")?.ToString()

            Case "pagination"
                _total = CInt(If(jo("total"), 0)) : _maxPage = CInt(If(jo("max_page"), 1))
                _currentReqTotal = _total   ' total SACEM pour cette requête
                pbProgress.Style = ProgressBarStyle.Blocks
                pbProgress.Maximum = Math.Max(_maxPage, 1) : pbProgress.Value = 0
                lblProgress.Text = $"{_total} oeuvre(s) · {_maxPage} page(s)"
                AddLog($"[INFO] {_total} oeuvre(s) — {_maxPage} page(s)")

            Case "oeuvres"
                Dim page As Integer = CInt(If(jo("page"), 1))
                Dim arr As JArray = TryCast(jo("oeuvres"), JArray)
                If arr Is Nothing Then Return
                For Each oe As JObject In arr
                    Dim tok As String = S(oe, "token")
                    Dim iswc As String = S(oe, "iswc").Trim()
                    Dim titre As String = S(oe, "titre").Trim().ToUpperInvariant()
                    Dim cle As String = If(iswc <> "", iswc, titre)
                    If _tokensConnus.Contains(cle) Then
                        _currentReqDbl += 1
                    Else
                        _tokensConnus.Add(cle)
                        _oeuvres.Add(oe)   ' stocké, pas encore dans le grid
                        _currentReqNouv += 1
                    End If
                Next
                pbProgress.Value = Math.Min(page, pbProgress.Maximum)
                lblProgress.Text = $"Page {page}/{_maxPage} · {_oeuvres.Count} oeuvre(s) récupérées…"
                AddLog($"[PAGE {page}/{_maxPage}] +{_currentReqNouv} nouv. / {_currentReqDbl} doublons · total {_oeuvres.Count}")
                AfficherStatLabel()

            Case "done"
                ' Si _enrichTotal > 0 → done du process --details-seulement → ignorer
                If _ignoreNextDone Then _ignoreNextDone = False : Return
                _statsParRequete.Add(Tuple.Create(_currentReqNom, _currentReqNouv, _currentReqDbl, _currentReqTotal))
                Dim nbNouv As Integer = _oeuvres.Count - _enrichOffset
                Dim n As Integer = CInt(If(jo("total_filtre_ipi"), 0))
                AddLog($"[OK] ""{_currentReqNom}"" : {_currentReqNouv} nouv. / {_currentReqDbl} doublons" &
                       If(n > 0, $" · {n} IPI", ""))
                AfficherStatLabel()

            Case "detail"
                Dim idxRaw As Integer = CInt(If(jo("index"), -1))
                Dim oed As JObject = TryCast(jo("oeuvre"), JObject)
                If idxRaw >= 0 AndAlso idxRaw < _oeuvres.Count AndAlso oed IsNot Nothing Then
                    _oeuvres(idxRaw) = oed
                    _detailsBuffer(idxRaw) = oed
                    _enrichDone += 1
                    ' Debug : vérifier si IPI présents
                    Dim auteurs As String = S(oed, "auteurs")
                    Dim hasIpi As Boolean = auteurs.Contains(" : ")
                    If _enrichDone <= 3 Then AddLog($"[DBG] idx={idxRaw} auteurs={auteurs.Replace(Chr(10), "/")} hasIpi={hasIpi}")
                    pbProgress.Value = Math.Min(_enrichDone, pbProgress.Maximum)
                    lblProgress.Text = $"Détails {_enrichDone}/{_enrichTotal}…"
                    If _enrichDone >= _enrichTotal AndAlso _enrichTotal > 0 Then
                        _modeDetails = False
                        _ignoreNextDone = True
                        lblProgress.Text = $"Mise à jour ({_detailsBuffer.Count} lignes)…"
                        For Each kv In _detailsBuffer
                            If kv.Key >= 0 AndAlso kv.Key < _dt.Rows.Count Then
                                EnrichirLigne(kv.Key, kv.Value)
                            End If
                        Next
                        _detailsBuffer.Clear()
                        ActiverCouleursDetail()
                        btnDetails.Text = "Détails OK"
                        AfficherStatLabel()
                        MajLblInfo()
                    End If
                End If

            Case "error"
                Dim m As String = jo("message")?.ToString()
                lblInfo.Text = "Erreur : " & m
                AddLog("[ERREUR] " & m)

        End Select
    End Sub

    Private Sub AjouterLigne(oe As JObject)
        Dim r As DataRow = _dt.NewRow()
        r("Sel") = False
        r("Titre") = S(oe, "titre")
        r("ISWC") = S(oe, "iswc")
        Dim typeVal As String = S(oe, "type")
        Dim genreVal As String = S(oe, "genre")
        r("Type") = If(typeVal <> "" AndAlso genreVal <> "" AndAlso typeVal <> genreVal,
                               typeVal & " / " & genreVal,
                               If(typeVal <> "", typeVal, genreVal))
        r("Duree") = S(oe, "duree")
        ' Colonnes rôles : déjà disponibles dès la phase 1 (paginatedData)
        r("A") = NomsPropres(S(oe, "auteurs"), "A")
        r("C") = NomsPropres(S(oe, "compositeurs"), "C")
        r("AR") = NomsPropres(S(oe, "auteurs") & Chr(10) & S(oe, "compositeurs"), "AR")
        r("AD") = NomsPropres(S(oe, "auteurs") & Chr(10) & S(oe, "compositeurs"), "AD")
        r("E") = NomsPropres(S(oe, "editeurs"), "")
        r("SE") = NomsPropres(S(oe, "sous_editeurs"), "")
        r("Interp") = NomsPropres(S(oe, "interpretes"), "")
        r("RE") = NomsPropres(S(oe, "realisateurs"), "")
        r("SousTitres") = S(oe, "sous_titres")
        r("Token") = S(oe, "token")
        r("IpiMatch") = CBool(If(oe("ipi_match")?.ToObject(Of Boolean)(), False))
        _dt.Rows.Add(r)
    End Sub

    Private Sub EnrichirLigne(idx As Integer, oe As JObject)
        If idx < 0 OrElse idx >= _dt.Rows.Count Then Return
        Dim r As DataRow = _dt.Rows(idx)
        ' Mettre à jour Type avec genre du détail (plus fiable que phase 1)
        Dim typeVal2 As String = S(oe, "type")
        Dim genreVal2 As String = S(oe, "genre")
        If typeVal2 <> "" OrElse genreVal2 <> "" Then
            r("Type") = If(typeVal2 <> "" AndAlso genreVal2 <> "" AndAlso typeVal2 <> genreVal2,
                           typeVal2 & " / " & genreVal2,
                           If(typeVal2 <> "", typeVal2, genreVal2))
        End If
        r("A") = NomsPropres(S(oe, "auteurs"), "A")
        r("C") = NomsPropres(S(oe, "compositeurs"), "C")
        r("AR") = NomsPropres(S(oe, "auteurs") & Chr(10) & S(oe, "compositeurs"), "AR")
        r("AD") = NomsPropres(S(oe, "auteurs") & Chr(10) & S(oe, "compositeurs"), "AD")
        r("E") = NomsPropres(S(oe, "editeurs"), "")
        r("SE") = NomsPropres(S(oe, "sous_editeurs"), "")
        r("Interp") = NomsPropres(S(oe, "interpretes"), "")
        r("RE") = NomsPropres(S(oe, "realisateurs"), "")
        r("SousTitres") = S(oe, "sous_titres")
        r("IpiMatch") = CBool(If(oe("ipi_match")?.ToObject(Of Boolean)(), False))
    End Sub

    ' Format entrée : "438009074 : NOM [CODE]" (une personne par \n)
    ' Format sortie : "438009074 : NOM\n..." (tag [xxx] retiré)
    Private Function NomsPropres(raw As String, roleCode As String) As String
        If String.IsNullOrEmpty(raw) Then Return ""
        Dim result As New List(Of String)()
        For Each entry In raw.Split(Chr(10))
            Dim e2 As String = entry.Trim()
            If String.IsNullOrEmpty(e2) Then Continue For
            If Not String.IsNullOrEmpty(roleCode) Then
                Dim isCA As Boolean = e2.ToUpper().Contains("[CA]") OrElse
                                      e2.ToUpper().Contains("[A+C]") OrElse
                                      e2.ToUpper().Contains("[AC]")
                Dim match As Boolean = e2.ToUpper().Contains($"[{roleCode.ToUpper()}]")
                If roleCode = "A" OrElse roleCode = "C" Then match = match OrElse isCA
                If Not match Then Continue For
            End If
            ' Tronquer au premier " | " (adresse/tel éditeur)
            Dim pipIdx As Integer = e2.IndexOf(" | ")
            If pipIdx > 0 Then e2 = e2.Substring(0, pipIdx).Trim()
            ' Retirer le tag [xxx] en fin de chaîne
            Dim ligne As String = System.Text.RegularExpressions.Regex.Replace(
                e2, "\s*\[.*?\]\s*$", "").Trim()
            If Not String.IsNullOrEmpty(ligne) Then result.Add(ligne)
        Next
        Return String.Join(Chr(10), result)
    End Function

    ' ══════════════════════════════════════════════════════════════════════
    ' DRAG LIGNES
    ' ══════════════════════════════════════════════════════════════════════
    Private _dragRowIndex As Integer = -1
    Private _dragStartPt As Point = Point.Empty

    Private Sub Dgv_RowDrag_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            Dim hit As DataGridView.HitTestInfo = dgvOeuvres.HitTest(e.X, e.Y)
            If hit.Type = DataGridViewHitTestType.Cell AndAlso hit.RowIndex >= 0 Then
                _dragRowIndex = hit.RowIndex
                _dragStartPt = New Point(e.X, e.Y)
            Else
                _dragRowIndex = -1
            End If
        End If
    End Sub

    Private Sub Dgv_RowDrag_MouseMove(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left AndAlso _dragRowIndex >= 0 Then
            ' Seuil minimum avant de déclencher le drag (évite les faux positifs)
            If Math.Abs(e.X - _dragStartPt.X) > SystemInformation.DragSize.Width OrElse
               Math.Abs(e.Y - _dragStartPt.Y) > SystemInformation.DragSize.Height Then
                dgvOeuvres.DoDragDrop(_dragRowIndex, DragDropEffects.Move)
                _dragRowIndex = -1
            End If
        End If
    End Sub

    Private Sub Dgv_RowDrag_DragOver(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.Move
        ' Surligner la ligne cible
        Dim pt As Point = dgvOeuvres.PointToClient(New Point(e.X, e.Y))
        Dim hit As DataGridView.HitTestInfo = dgvOeuvres.HitTest(pt.X, pt.Y)
        If hit.RowIndex >= 0 Then
            dgvOeuvres.Rows(hit.RowIndex).DefaultCellStyle.BackColor = Color.FromArgb(180, 210, 255)
        End If
    End Sub

    Private Sub Dgv_RowDrag_DragDrop(sender As Object, e As DragEventArgs)
        Dim pt As Point = dgvOeuvres.PointToClient(New Point(e.X, e.Y))
        Dim hit As DataGridView.HitTestInfo = dgvOeuvres.HitTest(pt.X, pt.Y)
        Dim targetIdx As Integer = hit.RowIndex

        If _dragRowIndex >= 0 AndAlso targetIdx >= 0 AndAlso
           Not (_dragRowIndex = targetIdx) AndAlso
           _dragRowIndex < _dt.Rows.Count AndAlso
           targetIdx < _dt.Rows.Count Then

            ' Copier la ligne source
            Dim src As DataRow = _dt.Rows(_dragRowIndex)
            Dim vals(src.ItemArray.Length - 1) As Object
            src.ItemArray.CopyTo(vals, 0)

            ' Supprimer et réinsérer à la position cible
            _dt.Rows.RemoveAt(_dragRowIndex)
            Dim newRow As DataRow = _dt.NewRow()
            newRow.ItemArray = vals
            _dt.Rows.InsertAt(newRow, If(targetIdx > _dragRowIndex, targetIdx, targetIdx))

            ' Synchroniser _oeuvres
            If _dragRowIndex < _oeuvres.Count AndAlso targetIdx <= _oeuvres.Count Then
                Dim oe As JObject = _oeuvres(_dragRowIndex)
                _oeuvres.RemoveAt(_dragRowIndex)
                _oeuvres.Insert(If(targetIdx > _dragRowIndex, targetIdx, targetIdx), oe)
            End If

            dgvOeuvres.ClearSelection()
            If targetIdx < dgvOeuvres.RowCount Then
                dgvOeuvres.Rows(If(targetIdx > _dragRowIndex, targetIdx, targetIdx)).Selected = True
            End If
        End If

        ' Réinitialiser couleurs alternatives
        For i As Integer = 0 To dgvOeuvres.RowCount - 1
            dgvOeuvres.Rows(i).DefaultCellStyle.BackColor = Color.Empty
        Next
        _dragRowIndex = -1
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' HANDLERS DGV
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub Dgv_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        If Not (dgvOeuvres.Columns(e.ColumnIndex).Name = "Titre") Then Return
        Dim drv As DataRowView = TryCast(_bs(e.RowIndex), DataRowView)
        If drv Is Nothing Then Return
        Dim token As String = CStr(If(drv.Row("Token"), ""))
        Dim titre As String = CStr(If(drv.Row("Titre"), ""))
        If String.IsNullOrEmpty(token) Then Return
        Dim url As String = "https://repertoire.sacem.fr/detail-oeuvre/" &
                            Uri.EscapeDataString(token) & "/" &
                            Uri.EscapeDataString(titre)
        Try
            Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show("Impossible d'ouvrir : " & ex.Message, "Erreur",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub Dgv_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        If dgvOeuvres.Columns(e.ColumnIndex).Name = "Titre" Then Return
        Dim drv As DataRowView = TryCast(_bs(e.RowIndex), DataRowView)
        If drv Is Nothing Then Return
        Dim token As String = CStr(If(drv.Row("Token"), ""))
        Dim idx As Integer = _oeuvres.FindIndex(Function(o) S(o, "token") = token)
        If idx < 0 Then Return
        Using f As New FicheSACEMOeuvreForm(_oeuvres(idx))
            f.ShowDialog(Me)
        End Using
    End Sub

    ' Coche uniquement sur clic dans la colonne Sel — bloque les autres colonnes
    Private Sub Dgv_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        If dgvOeuvres.Columns(e.ColumnIndex).Name <> "Sel" Then Return
        ' Valider immédiatement la coche
        dgvOeuvres.CommitEdit(DataGridViewDataErrorContexts.Commit)
    End Sub

    ' Ctrl+C → copier les cellules sélectionnées dans le presse-papiers
    Private Sub Dgv_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.C Then
            Dim sb As New System.Text.StringBuilder()
            For Each cell As DataGridViewCell In dgvOeuvres.SelectedCells
                If Not String.IsNullOrEmpty(CStr(If(cell.Value, ""))) Then
                    sb.AppendLine(CStr(cell.Value))
                End If
            Next
            If sb.Length > 0 Then
                Clipboard.SetText(sb.ToString().TrimEnd())
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Dgv_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
           dgvOeuvres.Columns(e.ColumnIndex).Name = "Titre" Then
            dgvOeuvres.Cursor = Cursors.Hand
        Else
            dgvOeuvres.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub Dgv_Scroll(sender As Object, e As ScrollEventArgs)
        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then PositionFilters()
    End Sub

    Private Sub Dgv_ColChanged(sender As Object, e As EventArgs)
        PositionFilters()
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' ÉTAT / LOGS
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub SetRunningState(running As Boolean)
        _running = running
        btnRechercher.Enabled = Not running
        btnAnnuler.Enabled = running
        btnSauvegarder.Enabled = Not running
        ' btnFiltrer et dgvOeuvres restent toujours actifs — grid non bloqué pendant recherche
        If running Then btnDetails.Enabled = False
        pnlProgress.Visible = running
        If Not running Then pbProgress.Style = ProgressBarStyle.Blocks
    End Sub

    ' Affiche toutes les requêtes du plan avec leur statut — appelée dès la saisie et à chaque mise à jour
    ' Lance N process --stat-seulement en parallèle, met à jour lblStats au fil des réponses
    ' Lance N process --stat-seulement en parallèle.
    ' Quand tous répondent → MessageBox Oui/Non. Oui = lance _pendingAction, Non = rien.
    Private Sub MajLblInfo()
        Dim tot As Integer = _dt.Rows.Count
        Dim vis As Integer = _dt.DefaultView.Count
        Dim sel As Integer = _dt.Rows.Cast(Of DataRow)().
                             Count(Function(r) CBool(If(r("Sel"), False)))
        Dim sb As New System.Text.StringBuilder()
        sb.Append($"{tot} ligne(s)")
        If vis < tot Then sb.Append($"  ·  {vis} visible(s)")
        If sel > 0 Then sb.Append($"  ·  {sel} sélectionnée(s)")
        lblInfo.Text = sb.ToString()
    End Sub

    Private Sub StatRecherche(pendingAction As Action)
        If _planRequetes.Count = 0 Then Return
        _statTotaux.Clear()
        Dim filtre As String = _pendingFiltre
        Dim attendus As Integer = _planRequetes.Count
        Dim reçus As Integer = 0

        For Each query As String In _planRequetes
            Dim q As String = query
            Task.Run(Sub()
                         Try
                             Dim psi As New ProcessStartInfo()
                             psi.FileName = "python"
                             psi.Arguments = $"""{Path.GetFullPath(ScriptPath)}"" --query ""{q}"" --filtre {filtre} --stat-seulement"""
                             psi.UseShellExecute = False
                             psi.RedirectStandardOutput = True
                             psi.StandardOutputEncoding = Encoding.UTF8
                             psi.CreateNoWindow = True
                             Dim proc As Process = Process.Start(psi)
                             Dim line As String = proc.StandardOutput.ReadLine()
                             proc.WaitForExit()
                             Dim total As Integer = 0
                             If line IsNot Nothing Then
                                 Try
                                     Dim jo As JObject = JObject.Parse(line)
                                     If jo("type")?.ToString() = "stat" Then total = CInt(If(jo("total"), 0))
                                 Catch : End Try
                             End If
                             If Me.IsHandleCreated Then
                                 Me.BeginInvoke(Sub()
                                                    _statTotaux(q) = total
                                                    reçus += 1
                                                    If reçus >= attendus Then
                                                        ' Tous reçus → construire message et afficher
                                                        Dim sb As New System.Text.StringBuilder()
                                                        Dim grandTotal As Integer = 0
                                                        For Each kv In _statTotaux
                                                            sb.AppendLine($"  {kv.Key} : {kv.Value} résultat(s)")
                                                            grandTotal += kv.Value
                                                        Next
                                                        If _planRequetes.Count > 1 Then
                                                            sb.AppendLine()
                                                            sb.AppendLine($"  Total estimé : ~{grandTotal}")
                                                        End If
                                                        sb.AppendLine()
                                                        sb.Append("Lancer la recherche ?")
                                                        Dim rep As DialogResult = MessageBox.Show(
                                    sb.ToString(), "Résultats SACEM",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                                        If rep = DialogResult.Yes Then
                                                            AfficherStatLabel()
                                                            pendingAction()
                                                        End If
                                                    End If
                                                End Sub)
                             End If
                         Catch : End Try
                     End Sub)
        Next
    End Sub

    ' Construit lblStats depuis _statTotaux + stats de recherche en cours
    Private Sub AfficherStatLabel()
        If _planRequetes.Count = 0 Then lblStats.Text = "" : Return
        Dim sb As New System.Text.StringBuilder()
        Dim totalDbl As Integer = If(_planRequetes.Count > 1, _statsParRequete.Sum(Function(t) t.Item3), 0)

        sb.Append($"  {_planRequetes.Count} req.  ║")

        For i As Integer = 0 To _planRequetes.Count - 1
            Dim nom As String = _planRequetes(i)
            sb.Append($"  {nom} : ")
            If i < _statsParRequete.Count Then
                sb.Append($"{_statsParRequete(i).Item4} ✓")
            ElseIf nom = _currentReqNom AndAlso _currentReqTotal > 0 Then
                sb.Append($"{_currentReqNouv}/{_currentReqTotal}…")
            ElseIf _statTotaux.ContainsKey(nom) Then
                sb.Append($"{_statTotaux(nom)}")
            Else
                sb.Append("?")
            End If
        Next

        If _planRequetes.Count > 1 AndAlso _statsParRequete.Count >= 2 Then
            sb.Append($"  ║  {totalDbl} commun(s)  ║  {_oeuvres.Count} total")
        ElseIf _planRequetes.Count > 1 AndAlso _oeuvres.Count > 0 Then
            sb.Append($"  ║  {_oeuvres.Count} total")
        End If

        lblStats.Text = sb.ToString()
    End Sub

    Private Sub OnProcessTermine()
        ' Si d'autres requêtes en attente → les enchaîner
        If Not _cancelRequested AndAlso _queryQueue.Count > 0 Then
            AddLog($"[NEXT] {_queryQueue.Count} requête(s) restante(s)…")
            LancerProchainerequete(_pendingFiltre, _pendingTitre)
            Return
        End If
        SetRunningState(False)
        If _cancelRequested Then
            AddLog($"[STOP] {_oeuvres.Count} oeuvre(s)")
        End If
        If _modeDetails Then
            ' Fin du process --details-seulement : rien à faire, le grid est déjà mis à jour
            _modeDetails = False
        ElseIf _oeuvres.Count > 0 Then
            ' Fin de la recherche : peupler le grid d'un coup
            lblProgress.Text = $"Chargement de {_oeuvres.Count} oeuvre(s)…"
            Application.DoEvents()
            _dt.BeginLoadData()
            For Each oe As JObject In _oeuvres
                AjouterLigne(oe)
            Next
            _dt.EndLoadData()
            btnDetails.Enabled = True
            btnDetails.Text = $"Détails ({_oeuvres.Count})"
            lblProgress.Text = $"{_oeuvres.Count} oeuvre(s) — cliquez Détails pour enrichir"
            AddLog($"[DETAIL] Toutes requêtes terminées — bouton Détails actif")
        End If
        Me.BeginInvoke(Sub() MajLblInfo())
    End Sub

    Private Sub BtnAnnuler_Click(sender As Object, e As EventArgs)
        _cancelRequested = True : KillProcess()
    End Sub

    Private Sub KillProcess()
        Try
            If _process IsNot Nothing AndAlso Not _process.HasExited Then _process.Kill()
        Catch
        End Try
    End Sub

    Private Sub AddLog(msg As String)
        txtLog.AppendText(If(txtLog.TextLength > 0, Environment.NewLine, "") &
                          $"[{DateTime.Now:HH:mm:ss}] {msg}")
        txtLog.SelectionStart = txtLog.TextLength
        txtLog.ScrollToCaret()
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' SÉLECTION + SAUVEGARDE
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub BtnSauvegarder_Click(sender As Object, e As EventArgs)
        Dim selection As New List(Of JObject)()
        For Each row As DataRow In _dt.Rows
            If CBool(If(row("Sel"), False)) Then
                Dim tok As String = CStr(If(row("Token"), ""))
                Dim idx As Integer = _oeuvres.FindIndex(Function(o) S(o, "token") = tok)
                If idx >= 0 Then selection.Add(_oeuvres(idx))
            End If
        Next
        If selection.Count = 0 Then selection.AddRange(_oeuvres)
        If selection.Count = 0 Then
            MessageBox.Show("Aucune oeuvre à sauvegarder.", "Sauvegarde",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        End If
        Try
            ExporterBDD(selection)
            lblInfo.Text = $"{selection.Count} oeuvre(s) sauvegardée(s) → {Path.GetFileName(BddPath)}"
            AddLog($"[SAVE] {selection.Count} oeuvre(s) → {Path.GetFullPath(BddPath)}")
        Catch ex As Exception
            MessageBox.Show("Erreur sauvegarde : " & ex.Message, "Erreur",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' EXPORT XLSX
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub ExporterBDD(oeuvres As List(Of JObject))
        Dim fi As New FileInfo(Path.GetFullPath(BddPath))
        Directory.CreateDirectory(fi.DirectoryName)
        Using pkg As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet = GetOrCreateSheet(pkg, "Oeuvres")
            EnsureHeaders(ws, {"Token", "Titre", "ISWC", "Type", "Durée",
                                "A", "C", "AR", "AD", "E", "SE", "IpiMatch", "DateExport"})
            For Each oe As JObject In oeuvres
                Dim tok As String = S(oe, "token")
                Dim r As Integer = FindOrAppend(ws, tok)
                ws.Cells(r, 1).Value = tok
                ws.Cells(r, 2).Value = S(oe, "titre")
                ws.Cells(r, 3).Value = S(oe, "iswc")
                ws.Cells(r, 4).Value = S(oe, "type")
                ws.Cells(r, 5).Value = S(oe, "duree")
                ws.Cells(r, 6).Value = NomsPropres(S(oe, "auteurs"), "A")
                ws.Cells(r, 7).Value = NomsPropres(S(oe, "compositeurs"), "C")
                ws.Cells(r, 8).Value = NomsPropres(S(oe, "auteurs") & Chr(10) & S(oe, "compositeurs"), "AR")
                ws.Cells(r, 9).Value = NomsPropres(S(oe, "auteurs") & Chr(10) & S(oe, "compositeurs"), "AD")
                ws.Cells(r, 10).Value = NomsPropres(S(oe, "editeurs"), "")
                ws.Cells(r, 11).Value = NomsPropres(S(oe, "sous_editeurs"), "")
                ws.Cells(r, 12).Value = If(CBool(If(oe("ipi_match")?.ToObject(Of Boolean)(), False)), "Oui", "")
                ws.Cells(r, 13).Value = DateTime.Now.ToString("yyyy-MM-dd")
            Next
            Dim wsI As ExcelWorksheet = GetOrCreateSheet(pkg, "Interpretes")
            EnsureHeaders(wsI, {"Token", "Interprètes"})
            For Each oe As JObject In oeuvres
                Dim tok As String = S(oe, "token")
                Dim r As Integer = FindOrAppend(wsI, tok)
                wsI.Cells(r, 1).Value = tok
                wsI.Cells(r, 2).Value = S(oe, "interpretes")
            Next
            pkg.SaveAs(fi)
        End Using
    End Sub

    Private Function GetOrCreateSheet(pkg As ExcelPackage, name As String) As ExcelWorksheet
        Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets(name)
        If ws Is Nothing Then ws = pkg.Workbook.Worksheets.Add(name)
        Return ws
    End Function

    Private Sub EnsureHeaders(ws As ExcelWorksheet, cols As String())
        If ws.Dimension Is Nothing OrElse ws.Cells(1, 1).Text = "" Then
            For i As Integer = 0 To cols.Length - 1
                ws.Cells(1, i + 1).Value = cols(i)
                ws.Cells(1, i + 1).Style.Font.Bold = True
            Next
        End If
    End Sub

    Private Function FindOrAppend(ws As ExcelWorksheet, key As String) As Integer
        If ws.Dimension IsNot Nothing Then
            For r As Integer = 2 To ws.Dimension.Rows
                If CStr(ws.Cells(r, 1).Value) = key Then Return r
            Next
        End If
        Return If(ws.Dimension IsNot Nothing, ws.Dimension.Rows + 1, 2)
    End Function

    ' ══════════════════════════════════════════════════════════════════════
    ' EXTRACTION RÔLES
    ' ══════════════════════════════════════════════════════════════════════
    Private Sub BtnExtraire_Click(sender As Object, e As EventArgs)
        ' Périmètre : toujours toutes les lignes du grid
        Dim rows As New List(Of DataRow)()
        For Each row As DataRow In _dt.Rows
            rows.Add(row)
        Next

        If rows.Count = 0 Then
            MessageBox.Show("Aucune ligne à extraire.", "Extraire rôles",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Colonnes rôle → label affiché
        Dim roleMap As New Dictionary(Of String, String) From {
            {"A", "A"}, {"C", "C"}, {"AR", "AR"}, {"AD", "AD"},
            {"E", "E"}, {"SE", "SE"}, {"Interp", "INT"}, {"RE", "RE"}
        }

        ' Termes de recherche actifs (hors watermark) — pour filtrer les entrées
        ' Si aucun terme → retourner tout
        Dim termesRecherche As New List(Of String)()
        If Not _rechWatermark AndAlso Not String.IsNullOrWhiteSpace(txtRecherche.Text) Then
            For Each grp In txtRecherche.Text.Split(New Char() {"*"c}, StringSplitOptions.RemoveEmptyEntries)
                For Each t2 In grp.Split(New Char() {"+"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t3 As String = t2.Trim()
                    If t3.StartsWith("!") Then t3 = t3.Substring(1).Trim()
                    If Not String.IsNullOrEmpty(t3) Then termesRecherche.Add(t3.ToUpper())
                Next
            Next
        End If

        ' Résultat : dédupliqué par (IPI+Nom), rôles regroupés
        ' Structure : clé = "IPI|NOM_UPPER" → (ipi, nom, List rôles)
        Dim vus As New Dictionary(Of String, Tuple(Of String, String, List(Of String)))()

        For Each row As DataRow In rows
            For Each kvp In roleMap
                Dim cellVal As String = CStr(If(row(kvp.Key), ""))
                If String.IsNullOrEmpty(cellVal) Then Continue For
                For Each entree In cellVal.Split(Chr(10))
                    Dim e2 As String = entree.Trim()
                    If String.IsNullOrEmpty(e2) Then Continue For

                    ' Parser "438009074 : NOM" d'abord
                    Dim ipi As String = ""
                    Dim nom As String = e2
                    Dim sep As Integer = e2.IndexOf(" : ")
                    If sep > 0 Then
                        ipi = e2.Substring(0, sep).Trim()
                        nom = e2.Substring(sep + 3).Trim()
                    End If

                    ' Filtrer par terme de recherche sur le nom (pas sur l'IPI)
                    If termesRecherche.Count > 0 Then
                        If Not termesRecherche.Any(Function(t4) nom.ToUpper().Contains(t4)) Then Continue For
                    End If

                    Dim cle As String = ipi & "|" & nom.ToUpper()
                    If Not vus.ContainsKey(cle) Then
                        vus(cle) = Tuple.Create(ipi, nom, New List(Of String)())
                    End If
                    If Not vus(cle).Item3.Contains(kvp.Value) Then
                        vus(cle).Item3.Add(kvp.Value)
                    End If
                Next
            Next
        Next

        Dim entries As New List(Of Tuple(Of String, String, List(Of String)))(vus.Values)
        Dim source As String = $"{rows.Count} ligne(s)"

        Dim cbAppliquer As Action(Of HashSet(Of String)) = Sub(cles As HashSet(Of String))
                                                               ' Décocher tout d'abord
                                                               For Each row As DataRow In _dt.Rows
                                                                   row("Sel") = False
                                                               Next
                                                               ' Cocher les lignes dont l'ISWC ou le titre est dans les clés
                                                               For Each row As DataRow In _dt.Rows
                                                                   Dim iswc As String = CStr(If(row("ISWC"), "")).Trim()
                                                                   Dim titreRaw As String = CStr(If(row("Titre"), "")).Trim()
                                                                   ' Tronquer au Chr(10) si sous-titre concaténé
                                                                   Dim nl As Integer = titreRaw.IndexOf(Chr(10))
                                                                   If nl > 0 Then titreRaw = titreRaw.Substring(0, nl).Trim()
                                                                   Dim titre As String = titreRaw.ToUpperInvariant()
                                                                   Dim cle As String = If(iswc <> "", iswc, titre)
                                                                   If cles.Contains(cle) Then row("Sel") = True
                                                               Next
                                                               MajLblInfo()
                                                           End Sub
        Using f As New FormExtraitRoles(entries, source, _oeuvres, cbAppliquer)
            f.ShowDialog(Me)
        End Using
    End Sub

    ' ══════════════════════════════════════════════════════════════════════
    ' HELPERS
    ' ══════════════════════════════════════════════════════════════════════
    Private Function S(jo As JObject, key As String) As String
        Dim t = jo(key)
        If t Is Nothing Then Return ""
        Return t.ToString().Trim()
    End Function

    Private Sub AddLbl(parent As Control, text As String, font As Font,
                       fore As Color, x As Integer, y As Integer)
        Dim l As New Label()
        l.Text = text : l.Font = font : l.ForeColor = fore
        l.Location = New Point(x, y) : l.AutoSize = True
        parent.Controls.Add(l)
    End Sub

    Private Function MkBtn(text As String, x As Integer, y As Integer,
                           w As Integer, bg As Color) As Button
        Dim b As New Button()
        b.Text = text : b.Location = New Point(x, y) : b.Size = New Size(w, 26)
        b.BackColor = bg : b.ForeColor = Color.White
        b.FlatStyle = FlatStyle.Flat : b.FlatAppearance.BorderSize = 0
        b.Font = New Font("Segoe UI", 8.5F) : b.Cursor = Cursors.Hand
        Return b
    End Function

End Class


' ════════════════════════════════════════════════════════════════════════════
' FicheSACEMOeuvreForm  —  Fiche lecture seule (double-clic)
' ════════════════════════════════════════════════════════════════════════════
Public Class FicheSACEMOeuvreForm
    Inherits Form

    Private Shared ReadOnly C_HEAD As Color = Color.FromArgb(20, 60, 140)
    Private Shared ReadOnly C_BG As Color = Color.FromArgb(230, 240, 255)
    Private Shared ReadOnly C_FIELD As Color = Color.FromArgb(215, 230, 252)

    Private ReadOnly _oe As JObject
    Private pnlMain As Panel

    Public Sub New(oe As JObject)
        _oe = oe : InitializeComponent() : RemplirFiche()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Fiche Oeuvre SACEM" : Me.Size = New Size(700, 680)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False : Me.BackColor = C_BG

        Dim ph As New Panel()
        ph.Dock = DockStyle.Top : ph.Height = 36 : ph.BackColor = C_HEAD
        Me.Controls.Add(ph)
        Dim lh As New Label()
        lh.Text = "RÉPERTOIRE PUBLIC SACEM  ·  Consultation (lecture seule)"
        lh.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        lh.ForeColor = Color.FromArgb(180, 210, 255)
        lh.Location = New Point(10, 9) : lh.AutoSize = True
        ph.Controls.Add(lh)

        pnlMain = New Panel()
        pnlMain.AutoScroll = True : pnlMain.Dock = DockStyle.Fill
        pnlMain.Padding = New Padding(12)
        Me.Controls.Add(pnlMain)

        Dim bc As New Button()
        bc.Text = "Fermer" : bc.Dock = DockStyle.Bottom : bc.Height = 30
        bc.BackColor = Color.FromArgb(60, 80, 140) : bc.ForeColor = Color.White
        bc.FlatStyle = FlatStyle.Flat : bc.FlatAppearance.BorderSize = 0
        bc.Font = New Font("Segoe UI", 9.0F) : bc.DialogResult = DialogResult.Cancel
        Me.Controls.Add(bc) : Me.CancelButton = bc
    End Sub

    Private Sub RemplirFiche()
        Dim y As Integer = 10
        y = Sect("OEUVRE", y)
        For Each kv In {("Titre", "titre"), ("ISWC", "iswc"), ("Token", "token"),
                        ("Genre", "genre"), ("Durée", "duree"), ("Type", "type")}
            y = Fld(kv.Item1, S(kv.Item2), y)
        Next
        Dim st As String = S("sous_titres")
        If Not String.IsNullOrEmpty(st) Then y = Fld("Sous-titres", st, y)

        Dim anyRole As Boolean = False
        For Each rm In {("A", "Auteur(s)"), ("C", "Compositeur(s)"),
                        ("AR", "Arrangeur(s)"), ("AD", "Adaptateur(s)")}
            Dim v As String = ParRole(rm.Item1)
            If Not String.IsNullOrEmpty(v) Then
                If Not anyRole Then y = Sect("AUTEURS / COMPOSITEURS", y) : anyRole = True
                y = FldML(rm.Item2, v, y)
            End If
        Next
        If Not anyRole Then
            Dim aut As String = S("auteurs") : Dim comp As String = S("compositeurs")
            If Not String.IsNullOrEmpty(aut) OrElse Not String.IsNullOrEmpty(comp) Then
                y = Sect("AUTEURS / COMPOSITEURS", y)
                If Not String.IsNullOrEmpty(aut) Then y = FldML("Auteurs", aut, y)
                If Not String.IsNullOrEmpty(comp) AndAlso Not (comp = aut) Then y = FldML("Compositeurs", comp, y)
            End If
        End If

        Dim ed As String = S("editeurs") : Dim se As String = S("sous_editeurs")
        If Not String.IsNullOrEmpty(ed) OrElse Not String.IsNullOrEmpty(se) Then
            y = Sect("ÉDITEURS", y)
            If Not String.IsNullOrEmpty(ed) Then y = FldML("Éditeur(s)", ed, y)
            If Not String.IsNullOrEmpty(se) Then y = FldML("Sous-éditeur(s)", se, y)
        End If

        Dim interps As String = S("interpretes")
        If Not String.IsNullOrEmpty(interps) Then
            y = Sect("INTERPRÈTES", y)
            y = FldML("Noms", interps, y)
        End If

        pnlMain.AutoScrollMinSize = New Size(0, y + 20)
    End Sub

    Private Function ParRole(roleCode As String) As String
        Dim result As New List(Of String)()
        For Each champ In {"auteurs", "compositeurs"}
            Dim val As String = S(champ)
            If String.IsNullOrEmpty(val) Then Continue For
            For Each entry In val.Split(Chr(10))
                Dim e2 As String = entry.Trim()
                If String.IsNullOrEmpty(e2) Then Continue For
                Dim isCA As Boolean = e2.ToUpper().Contains("[CA]") OrElse
                                      e2.ToUpper().Contains("[A+C]") OrElse
                                      e2.ToUpper().Contains("[AC]")
                Dim match As Boolean = e2.ToUpper().Contains($"[{roleCode.ToUpper()}]")
                If roleCode = "A" OrElse roleCode = "C" Then match = match OrElse isCA
                If Not match Then Continue For
                Dim ligne As String = System.Text.RegularExpressions.Regex.Replace(
                    e2, "\s*\[.*?\]\s*$", "").Trim()
                If Not String.IsNullOrEmpty(ligne) Then result.Add(ligne)
            Next
        Next
        Return String.Join(Chr(10), result)
    End Function

    Private Function Sect(titre As String, y As Integer) As Integer
        Dim lbl As New Label()
        lbl.Text = titre : lbl.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        lbl.ForeColor = C_HEAD : lbl.Location = New Point(10, y) : lbl.AutoSize = True
        pnlMain.Controls.Add(lbl)
        Dim sep As New Panel()
        sep.Location = New Point(10, y + 20) : sep.Size = New Size(640, 1)
        sep.BackColor = Color.FromArgb(160, 190, 230)
        pnlMain.Controls.Add(sep)
        Return y + 28
    End Function

    ' Champ simple ligne
    Private Function Fld(label As String, value As String, y As Integer) As Integer
        If String.IsNullOrEmpty(value) Then Return y
        Dim lbl As New Label()
        lbl.Text = label & ":" : lbl.Font = New Font("Segoe UI", 7.5F)
        lbl.ForeColor = Color.FromArgb(110, 130, 160)
        lbl.Location = New Point(10, y) : lbl.Size = New Size(115, 14)
        pnlMain.Controls.Add(lbl)
        Dim tb As New TextBox()
        tb.Text = value : tb.Location = New Point(130, y - 2)
        tb.Size = New Size(525, 22) : tb.ReadOnly = True
        tb.BorderStyle = BorderStyle.None : tb.BackColor = C_FIELD
        tb.Font = New Font("Segoe UI", 9.0F) : tb.ForeColor = Color.FromArgb(20, 40, 90)
        pnlMain.Controls.Add(tb)
        Return y + 26
    End Function

    ' Champ multi-ligne (liste IPI : NOM)
    Private Function FldML(label As String, value As String, y As Integer) As Integer
        If String.IsNullOrEmpty(value) Then Return y
        Dim lines As Integer = value.Split(Chr(10)).Length
        Dim h As Integer = Math.Max(22, lines * 18)
        Dim lbl As New Label()
        lbl.Text = label & ":" : lbl.Font = New Font("Segoe UI", 7.5F)
        lbl.ForeColor = Color.FromArgb(110, 130, 160)
        lbl.Location = New Point(10, y) : lbl.Size = New Size(115, 14)
        pnlMain.Controls.Add(lbl)
        Dim tb As New TextBox()
        tb.Text = value.Replace(Chr(10), Environment.NewLine)
        tb.Location = New Point(130, y - 2)
        tb.Size = New Size(525, h + 4) : tb.ReadOnly = True
        tb.Multiline = True : tb.ScrollBars = ScrollBars.None
        tb.BorderStyle = BorderStyle.None : tb.BackColor = C_FIELD
        tb.Font = New Font("Segoe UI", 9.0F) : tb.ForeColor = Color.FromArgb(20, 40, 90)
        pnlMain.Controls.Add(tb)
        Return y + h + 8
    End Function

    Private Function S(key As String) As String
        Dim t = _oe(key)
        If t Is Nothing Then Return ""
        Return t.ToString().Trim()
    End Function

End Class

' ════════════════════════════════════════════════════════════════════════════
' ════════════════════════════════════════════════════════════════════════════
' ════════════════════════════════════════════════════════════════════════════
' ════════════════════════════════════════════════════════════════════════════
' ════════════════════════════════════════════════════════════════════════════
' ════════════════════════════════════════════════════════════════════════════
' FormExtraitRoles  —  ListView checkboxes + panneau titres copiable
' ════════════════════════════════════════════════════════════════════════════
Public Class FormExtraitRoles
    Inherits Form

    Private _lv As ListView
    Private _txtInfo As TextBox
    Private _oeuvres As List(Of JObject)
    Private _entries As List(Of Tuple(Of String, String, List(Of String)))
    Private _cbAppliquer As Action(Of HashSet(Of String))

    Public Sub New(entries As List(Of Tuple(Of String, String, List(Of String))), source As String, oeuvres As List(Of JObject), Optional cbAppliquer As Action(Of HashSet(Of String)) = Nothing)
        Me.Text = "Extrait des rôles"
        Me.Size = New Size(1060, 560)
        Me.MinimumSize = New Size(700, 350)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(230, 240, 255)
        _oeuvres = oeuvres
        _entries = entries
        _cbAppliquer = cbAppliquer

        ' ── Panneau droit : titres copiables ──────────────────────────────
        Dim pRight As New Panel()
        pRight.Dock = DockStyle.Right
        pRight.Width = 460
        pRight.BackColor = Color.FromArgb(235, 243, 255)

        _txtInfo = New TextBox()
        _txtInfo.Dock = DockStyle.Fill
        _txtInfo.Multiline = True
        _txtInfo.ReadOnly = True
        _txtInfo.ScrollBars = ScrollBars.Both
        _txtInfo.Font = New Font("Consolas", 8.5F)
        _txtInfo.ForeColor = Color.FromArgb(20, 40, 90)
        _txtInfo.BackColor = Color.FromArgb(240, 248, 255)
        _txtInfo.BorderStyle = BorderStyle.None
        _txtInfo.WordWrap = False
        pRight.Controls.Add(_txtInfo)

        Dim lblRight As New Label()
        lblRight.Text = "Cliquez un nom pour voir ses titres"
        lblRight.Font = New Font("Segoe UI", 8.0F, FontStyle.Italic)
        lblRight.ForeColor = Color.FromArgb(100, 120, 160)
        lblRight.Dock = DockStyle.Top
        lblRight.Height = 22
        lblRight.Padding = New Padding(4, 4, 0, 0)
        pRight.Controls.Add(lblRight)
        Me.Controls.Add(pRight)

        ' Séparateur
        Dim sep As New Panel()
        sep.Dock = DockStyle.Right
        sep.Width = 2
        sep.BackColor = Color.FromArgb(180, 200, 240)
        Me.Controls.Add(sep)

        ' ── ListView gauche avec checkboxes ───────────────────────────────
        _lv = New ListView()
        _lv.Dock = DockStyle.Fill
        _lv.View = View.Details
        _lv.CheckBoxes = True
        _lv.FullRowSelect = True
        _lv.GridLines = True
        _lv.Font = New Font("Consolas", 9.0F)
        _lv.BackColor = Color.FromArgb(240, 246, 255)
        _lv.Columns.Add("Rôle(s)", 90)
        _lv.Columns.Add("IPI", 115)
        _lv.Columns.Add("Nom", 280)
        For Each e In entries
            Dim item As New ListViewItem(String.Join(",", e.Item3))
            item.SubItems.Add(e.Item1)
            item.SubItems.Add(e.Item2)
            _lv.Items.Add(item)
        Next
        AddHandler _lv.ItemSelectionChanged, AddressOf Lv_SelectionChanged
        Me.Controls.Add(_lv)

        ' ── Bandeau ────────────────────────────────────────────────────────
        Dim ph As New Panel() : ph.Dock = DockStyle.Top : ph.Height = 36
        ph.BackColor = Color.FromArgb(20, 60, 140)
        Dim lh As New Label()
        lh.Text = "EXTRAIT RÔLES  ·  " & source
        lh.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        lh.ForeColor = Color.FromArgb(180, 210, 255)
        lh.Location = New Point(10, 9) : lh.AutoSize = True
        ph.Controls.Add(lh)
        Me.Controls.Add(ph)

        ' ── Barre bas ──────────────────────────────────────────────────────
        Dim pBot As New Panel() : pBot.Dock = DockStyle.Bottom : pBot.Height = 38
        pBot.BackColor = Color.FromArgb(215, 228, 250)

        Dim lblNb As New Label()
        lblNb.Text = $"{entries.Count} entrée(s)"
        lblNb.Font = New Font("Segoe UI", 8.0F)
        lblNb.ForeColor = Color.FromArgb(60, 80, 120)
        lblNb.Location = New Point(10, 11) : lblNb.AutoSize = True
        pBot.Controls.Add(lblNb)

        Dim btnTout As New Button()
        btnTout.Text = "Tout"
        btnTout.Location = New Point(130, 6) : btnTout.Size = New Size(60, 26)
        btnTout.BackColor = Color.FromArgb(80, 100, 140) : btnTout.ForeColor = Color.White
        btnTout.FlatStyle = FlatStyle.Flat : btnTout.FlatAppearance.BorderSize = 0
        btnTout.Font = New Font("Segoe UI", 8.5F) : btnTout.Cursor = Cursors.Hand
        AddHandler btnTout.Click, Sub(s, ev)
                                      Dim tousCoches As Boolean = _lv.CheckedItems.Count = _lv.Items.Count
                                      For Each item As ListViewItem In _lv.Items
                                          item.Checked = Not tousCoches
                                      Next
                                  End Sub
        pBot.Controls.Add(btnTout)

        Dim btnSel As New Button()
        btnSel.Text = "Titres cochés"
        btnSel.Location = New Point(200, 6) : btnSel.Size = New Size(120, 26)
        btnSel.BackColor = Color.FromArgb(20, 110, 60) : btnSel.ForeColor = Color.White
        btnSel.FlatStyle = FlatStyle.Flat : btnSel.FlatAppearance.BorderSize = 0
        btnSel.Font = New Font("Segoe UI", 8.5F) : btnSel.Cursor = Cursors.Hand
        AddHandler btnSel.Click, AddressOf BtnSel_Click
        pBot.Controls.Add(btnSel)

        Dim btnCopier As New Button()
        btnCopier.Text = "Copier titres"
        btnCopier.Location = New Point(330, 6) : btnCopier.Size = New Size(110, 26)
        btnCopier.BackColor = Color.FromArgb(40, 80, 160) : btnCopier.ForeColor = Color.White
        btnCopier.FlatStyle = FlatStyle.Flat : btnCopier.FlatAppearance.BorderSize = 0
        btnCopier.Font = New Font("Segoe UI", 8.5F) : btnCopier.Cursor = Cursors.Hand
        AddHandler btnCopier.Click, Sub(s, ev)
                                        If _txtInfo.Text.Trim() = "" Then Return
                                        Clipboard.SetText(_txtInfo.Text)
                                    End Sub
        pBot.Controls.Add(btnCopier)

        Dim btnAppliquer As New Button()
        btnAppliquer.Text = "Appliquer selection"
        btnAppliquer.Location = New Point(450, 6) : btnAppliquer.Size = New Size(150, 26)
        btnAppliquer.BackColor = Color.FromArgb(140, 60, 20) : btnAppliquer.ForeColor = Color.White
        btnAppliquer.FlatStyle = FlatStyle.Flat : btnAppliquer.FlatAppearance.BorderSize = 0
        btnAppliquer.Font = New Font("Segoe UI", 8.5F) : btnAppliquer.Cursor = Cursors.Hand
        btnAppliquer.Enabled = (_cbAppliquer IsNot Nothing)
        AddHandler btnAppliquer.Click, AddressOf BtnAppliquer_Click
        pBot.Controls.Add(btnAppliquer)

        Dim btnF As New Button()
        btnF.Text = "Fermer" : btnF.Dock = DockStyle.Right : btnF.Width = 80
        btnF.BackColor = Color.FromArgb(80, 80, 110) : btnF.ForeColor = Color.White
        btnF.FlatStyle = FlatStyle.Flat : btnF.FlatAppearance.BorderSize = 0
        btnF.DialogResult = DialogResult.Cancel
        pBot.Controls.Add(btnF)
        Me.CancelButton = btnF
        Me.Controls.Add(pBot)
    End Sub

    Private Function NomsCoches() As HashSet(Of String)
        Dim noms As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each item As ListViewItem In _lv.CheckedItems
            Dim idx As Integer = _lv.Items.IndexOf(item)
            If idx >= 0 AndAlso idx < _entries.Count Then
                noms.Add(_entries(idx).Item2.Trim())
            End If
        Next
        Return noms
    End Function

    Private Sub Lv_SelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs)
        If Not e.IsSelected Then Return
        Dim idx As Integer = e.ItemIndex
        If idx < 0 OrElse idx >= _entries.Count Then Return
        Dim nom As String = _entries(idx).Item2.Trim()
        If String.IsNullOrEmpty(nom) Then Return
        Dim nomUp As String = nom.ToUpperInvariant()

        Dim titres = _oeuvres.Where(Function(oe)
                                        For Each champ In {"auteurs", "compositeurs", "editeurs", "sous_editeurs", "interpretes", "realisateurs"}
                                            If If(oe(champ)?.ToString(), "").ToUpperInvariant().Contains(nomUp) Then Return True
                                        Next
                                        Return False
                                    End Function).ToList()

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($"=== {titres.Count} titre(s) — {nom} ===")
        For Each oe As JObject In titres
            sb.AppendLine(If(oe("titre")?.ToString(), ""))
        Next
        _txtInfo.Text = sb.ToString()
        _txtInfo.SelectionStart = 0
        _txtInfo.ScrollToCaret()
    End Sub

    Private Sub BtnSel_Click(sender As Object, e As EventArgs)
        Dim noms = NomsCoches()
        If noms.Count = 0 Then _txtInfo.Text = "(aucun nom coché)" : Return
        Dim vus As New HashSet(Of String)()
        Dim sb As New System.Text.StringBuilder()
        Dim count As Integer = 0
        For Each oe As JObject In _oeuvres
            Dim match As Boolean = False
            Dim ad As String = If(oe("ayants_droits")?.ToString(), "").ToUpperInvariant()
            For Each nom In noms
                If ad.Contains(nom.ToUpperInvariant()) Then match = True : Exit For
            Next
            If Not match Then
                For Each champ In {"auteurs", "compositeurs", "editeurs", "sous_editeurs", "interpretes", "realisateurs"}
                    Dim val As String = If(oe(champ)?.ToString(), "").ToUpperInvariant()
                    For Each nom In noms
                        If val.Contains(nom.ToUpperInvariant()) Then match = True : Exit For
                    Next
                    If match Then Exit For
                Next
            End If
            If Not match Then Continue For
            Dim cle As String = If(oe("iswc")?.ToString().Trim() <> "", oe("iswc").ToString().Trim(), oe("titre")?.ToString())
            If vus.Contains(cle) Then Continue For
            vus.Add(cle)
            sb.AppendLine(If(oe("titre")?.ToString(), ""))
            count += 1
        Next
        _txtInfo.Text = $"=== {count} titre(s) pour : {String.Join(", ", noms)} ===" &
                        Environment.NewLine & New String("-"c, 50) & Environment.NewLine & sb.ToString()
        _txtInfo.SelectionStart = 0
        _txtInfo.ScrollToCaret()
    End Sub

    Private Sub BtnAppliquer_Click(sender As Object, e As EventArgs)
        Dim noms = NomsCoches()
        If noms.Count = 0 Then
            MessageBox.Show("Cochez au moins un nom.", "Appliquer", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim cles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each oe As JObject In _oeuvres
            Dim match As Boolean = False
            Dim ad As String = If(oe("ayants_droits")?.ToString(), "").ToUpperInvariant()
            For Each nom In noms
                If ad.Contains(nom.ToUpperInvariant()) Then match = True : Exit For
            Next
            If Not match Then
                For Each champ In {"auteurs", "compositeurs", "editeurs", "sous_editeurs", "interpretes", "realisateurs"}
                    Dim val As String = If(oe(champ)?.ToString(), "").ToUpperInvariant()
                    For Each nom In noms
                        If val.Contains(nom.ToUpperInvariant()) Then match = True : Exit For
                    Next
                    If match Then Exit For
                Next
            End If
            If Not match Then Continue For
            Dim iswc As String = If(oe("iswc")?.ToString().Trim(), "")
            Dim titre As String = If(oe("titre")?.ToString().Trim().ToUpperInvariant(), "")
            cles.Add(If(iswc <> "", iswc, titre))
        Next
        _cbAppliquer(cles)
        Me.Close()
    End Sub

End Class





' ════════════════════════════════════════════════════════════════════════════
' FormTitresParNom  —  Liste des titres du grid associés à un nom
' ════════════════════════════════════════════════════════════════════════════
Public Class FormTitresParNom
    Inherits Form

    Public Sub New(nom As String, oeuvres As List(Of JObject))
        Me.Text = $"Titres — {nom}"
        Me.Size = New Size(760, 480)
        Me.MinimumSize = New Size(500, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(230, 240, 255)

        ' Bandeau
        Dim ph As New Panel() : ph.Dock = DockStyle.Top : ph.Height = 36
        ph.BackColor = Color.FromArgb(20, 60, 140)
        Dim lh As New Label()
        lh.Text = $"{oeuvres.Count} titre(s) associé(s) à  {nom}"
        lh.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        lh.ForeColor = Color.FromArgb(180, 210, 255)
        lh.Location = New Point(10, 9) : lh.AutoSize = True
        ph.Controls.Add(lh)

        ' Bouton fermer bas
        Dim pBot As New Panel() : pBot.Dock = DockStyle.Bottom : pBot.Height = 36
        pBot.BackColor = Color.FromArgb(215, 228, 250)
        Dim btnF As New Button()
        btnF.Text = "Fermer" : btnF.Dock = DockStyle.Right : btnF.Width = 80
        btnF.BackColor = Color.FromArgb(80, 80, 110) : btnF.ForeColor = Color.White
        btnF.FlatStyle = FlatStyle.Flat : btnF.FlatAppearance.BorderSize = 0
        btnF.DialogResult = DialogResult.Cancel
        pBot.Controls.Add(btnF)
        Me.CancelButton = btnF

        ' ListView
        Dim lv As New ListView()
        lv.Dock = DockStyle.Fill
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        lv.Font = New Font("Segoe UI", 8.5F)
        lv.BackColor = Color.FromArgb(240, 246, 255)
        lv.Columns.Add("Titre", 320)
        lv.Columns.Add("ISWC", 140)
        lv.Columns.Add("Type", 100)
        lv.Columns.Add("Durée", 60)
        lv.Columns.Add("Rôle", 100)

        For Each oe As JObject In oeuvres
            Dim titre As String = If(oe("titre")?.ToString(), "")
            Dim iswc As String = If(oe("iswc")?.ToString(), "")
            Dim type_ As String = If(oe("type")?.ToString(), "")
            Dim duree As String = If(oe("duree")?.ToString(), "")
            ' Déterminer le rôle de ce nom dans cette oeuvre
            Dim nomUp As String = nom.ToUpperInvariant()
            Dim roles As New List(Of String)()
            If If(oe("auteurs")?.ToString(), "").ToUpperInvariant().Contains(nomUp) Then roles.Add("A")
            If If(oe("compositeurs")?.ToString(), "").ToUpperInvariant().Contains(nomUp) Then roles.Add("C")
            If If(oe("editeurs")?.ToString(), "").ToUpperInvariant().Contains(nomUp) Then roles.Add("E")
            If If(oe("sous_editeurs")?.ToString(), "").ToUpperInvariant().Contains(nomUp) Then roles.Add("SE")
            If If(oe("interpretes")?.ToString(), "").ToUpperInvariant().Contains(nomUp) Then roles.Add("INT")
            Dim role As String = If(roles.Count > 0, String.Join(",", roles), "?")
            Dim item As New ListViewItem(titre)
            item.SubItems.Add(iswc)
            item.SubItems.Add(type_)
            item.SubItems.Add(duree)
            item.SubItems.Add(role)
            lv.Items.Add(item)
        Next

        ' Ordre d'ajout : Fill avant Bottom et Top
        Me.Controls.Add(lv)
        Me.Controls.Add(pBot)
        Me.Controls.Add(ph)
    End Sub

End Class
