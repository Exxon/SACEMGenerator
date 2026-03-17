Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data

Public Enum ChoixEditeurType
    PartInedite
    EAC
    ChoisirEditeur
End Enum

Public Class FormChoixEditeur
    Inherits Form

    Public Property Choix As ChoixEditeurType = ChoixEditeurType.PartInedite
    Public Property EditeurId As String = ""
    Public Property RoleChoisi As String = ""

    Private _dtMor As DataTable
    Private _rolesDisponibles As List(Of String)

    Private rbInedite As RadioButton
    Private rbEAC As RadioButton
    Private rbChoisir As RadioButton
    Private cbEditeurs As ComboBox
    Private cbRole As ComboBox
    Private btnOK As Button
    Private btnAnnuler As Button

    Public Sub New(data As DataTable, Optional roles As List(Of String) = Nothing)
        _dtMor = data
        _rolesDisponibles = If(roles IsNot Nothing AndAlso roles.Count > 0, roles, New List(Of String)())
        InitForm()
    End Sub

    Private Sub InitForm()
        Dim hasRoleChoice As Boolean = (_rolesDisponibles.Count > 1)
        Me.Text = "Association éditeur"
        Me.Size = New Size(400, If(hasRoleChoice, 310, 260))
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(245, 247, 250)

        Dim y As Integer = 14

        If hasRoleChoice Then
            Dim lblRole As New Label()
            lblRole.Text = "Rôle :"
            lblRole.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            lblRole.Location = New Point(16, y)
            lblRole.Size = New Size(60, 22)
            Me.Controls.Add(lblRole)

            cbRole = New ComboBox()
            cbRole.Font = New Font("Segoe UI", 9)
            cbRole.Location = New Point(80, y)
            cbRole.Size = New Size(120, 24)
            cbRole.DropDownStyle = ComboBoxStyle.DropDownList
            For Each r As String In _rolesDisponibles
                cbRole.Items.Add(r)
            Next
            cbRole.SelectedIndex = 0
            Me.Controls.Add(cbRole)
            y += 46
        End If

        Dim lbl As New Label()
        lbl.Text = "Cet ayant droit n'a pas d'éditeur associé."
        lbl.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lbl.Location = New Point(16, y)
        lbl.Size = New Size(360, 18)
        Me.Controls.Add(lbl)
        y += 22

        Dim lbl2 As New Label()
        lbl2.Text = "Choisissez une option :"
        lbl2.Font = New Font("Segoe UI", 9)
        lbl2.Location = New Point(16, y)
        lbl2.Size = New Size(360, 18)
        Me.Controls.Add(lbl2)
        y += 26

        rbInedite = New RadioButton()
        rbInedite.Text = "Part inédite (aucun éditeur)"
        rbInedite.Font = New Font("Segoe UI", 9)
        rbInedite.Location = New Point(24, y)
        rbInedite.Size = New Size(340, 22)
        rbInedite.Checked = True
        Me.Controls.Add(rbInedite)
        y += 28

        rbEAC = New RadioButton()
        rbEAC.Text = "Éditeur À Compte d'Auteur (EAC)"
        rbEAC.Font = New Font("Segoe UI", 9)
        rbEAC.Location = New Point(24, y)
        rbEAC.Size = New Size(340, 22)
        Me.Controls.Add(rbEAC)
        y += 28

        rbChoisir = New RadioButton()
        rbChoisir.Text = "Choisir un éditeur :"
        rbChoisir.Font = New Font("Segoe UI", 9)
        rbChoisir.Location = New Point(24, y)
        rbChoisir.Size = New Size(200, 22)
        Me.Controls.Add(rbChoisir)
        y += 28

        cbEditeurs = New ComboBox()
        cbEditeurs.Font = New Font("Segoe UI", 9)
        cbEditeurs.Location = New Point(44, y)
        cbEditeurs.Size = New Size(320, 24)
        cbEditeurs.DropDownStyle = ComboBoxStyle.DropDownList
        cbEditeurs.Enabled = False
        Me.Controls.Add(cbEditeurs)

        If _dtMor IsNot Nothing Then
            For Each row As DataRow In _dtMor.Rows
                Dim designation As String = If(row.Table.Columns.Contains("Designation"), row("Designation").ToString().Trim(), "")
                Dim id As String = If(row.Table.Columns.Contains("Id"), row("Id").ToString().Trim(), "")
                If Not String.IsNullOrEmpty(designation) Then
                    cbEditeurs.Items.Add(New EditeurItem(id, designation))
                End If
            Next
            If cbEditeurs.Items.Count > 0 Then cbEditeurs.SelectedIndex = 0
        End If

        AddHandler rbChoisir.CheckedChanged, AddressOf RbChoisir_CheckedChanged

        y += 36

        btnOK = New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(80, 28)
        btnOK.Location = New Point(210, y)
        btnOK.BackColor = Color.FromArgb(70, 130, 180)
        btnOK.ForeColor = Color.White
        btnOK.FlatStyle = FlatStyle.Flat
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        Me.Controls.Add(btnOK)

        btnAnnuler = New Button()
        btnAnnuler.Text = "Annuler"
        btnAnnuler.Size = New Size(80, 28)
        btnAnnuler.Location = New Point(300, y)
        btnAnnuler.FlatStyle = FlatStyle.Flat
        AddHandler btnAnnuler.Click, AddressOf BtnAnnuler_Click
        Me.Controls.Add(btnAnnuler)

        Me.AcceptButton = btnOK
        Me.CancelButton = btnAnnuler
    End Sub

    Private Sub RbChoisir_CheckedChanged(sender As Object, e As EventArgs)
        cbEditeurs.Enabled = rbChoisir.Checked
    End Sub

    Private Sub BtnAnnuler_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        If cbRole IsNot Nothing AndAlso cbRole.SelectedItem IsNot Nothing Then
            RoleChoisi = cbRole.SelectedItem.ToString()
        End If
        If rbInedite.Checked Then
            Choix = ChoixEditeurType.PartInedite
        ElseIf rbEAC.Checked Then
            Choix = ChoixEditeurType.EAC
        ElseIf rbChoisir.Checked Then
            Choix = ChoixEditeurType.ChoisirEditeur
            Dim item As EditeurItem = TryCast(cbEditeurs.SelectedItem, EditeurItem)
            If item Is Nothing Then
                MessageBox.Show("Veuillez sélectionner un éditeur.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            EditeurId = item.Id
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Class EditeurItem
        Public ReadOnly Id As String
        Public ReadOnly Designation As String
        Public Sub New(id As String, designation As String)
            Me.Id = id
            Me.Designation = designation
        End Sub
        Public Overrides Function ToString() As String
            Return Designation
        End Function
    End Class

End Class
