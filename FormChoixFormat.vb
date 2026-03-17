'Imports System.Windows.Forms
Imports System.Drawing
Imports Microsoft.Win32

Public Enum FormatSortie
    DocxUniquement
    PdfUniquement
    DocxEtPdf
End Enum

Public Class FormChoixFormat
    Inherits Form

    Public Property FormatChoisi As FormatSortie = FormatSortie.DocxUniquement

    Private rbDocx As RadioButton
    Private rbPdf As RadioButton
    Private rbLesDeux As RadioButton
    Private cbMemoriser As CheckBox
    Private btnOK As Button
    Private btnAnnuler As Button

    Private Const REG_KEY As String = "SOFTWARE\SACEMGenerator"
    Private Const REG_FORMAT As String = "FormatSortieContrat"
    Private Const REG_MEMORISE As String = "FormatSortieMemoriser"

    Public Sub New()
        InitForm()
        ChargerPreference()
    End Sub

    Private Sub InitForm()
        Me.Text = "Format de sortie"
        Me.Size = New Size(320, 230)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(245, 247, 250)

        Dim lbl As New Label()
        lbl.Text = "Choisissez le format de génération :"
        lbl.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lbl.Location = New Point(16, 16)
        lbl.Size = New Size(280, 18)
        Me.Controls.Add(lbl)

        rbDocx = New RadioButton()
        rbDocx.Text = "DOCX uniquement"
        rbDocx.Font = New Font("Segoe UI", 9)
        rbDocx.Location = New Point(24, 44)
        rbDocx.Size = New Size(260, 22)
        rbDocx.Checked = True
        Me.Controls.Add(rbDocx)

        rbPdf = New RadioButton()
        rbPdf.Text = "PDF uniquement"
        rbPdf.Font = New Font("Segoe UI", 9)
        rbPdf.Location = New Point(24, 70)
        rbPdf.Size = New Size(260, 22)
        Me.Controls.Add(rbPdf)

        rbLesDeux = New RadioButton()
        rbLesDeux.Text = "DOCX + PDF"
        rbLesDeux.Font = New Font("Segoe UI", 9)
        rbLesDeux.Location = New Point(24, 96)
        rbLesDeux.Size = New Size(260, 22)
        Me.Controls.Add(rbLesDeux)

        cbMemoriser = New CheckBox()
        cbMemoriser.Text = "Mémoriser ce choix"
        cbMemoriser.Font = New Font("Segoe UI", 9)
        cbMemoriser.Location = New Point(24, 128)
        cbMemoriser.Size = New Size(260, 20)
        Me.Controls.Add(cbMemoriser)

        btnOK = New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(80, 28)
        btnOK.Location = New Point(120, 158)
        btnOK.BackColor = Color.FromArgb(70, 130, 180)
        btnOK.ForeColor = Color.White
        btnOK.FlatStyle = FlatStyle.Flat
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        Me.Controls.Add(btnOK)

        btnAnnuler = New Button()
        btnAnnuler.Text = "Annuler"
        btnAnnuler.Size = New Size(80, 28)
        btnAnnuler.Location = New Point(210, 158)
        btnAnnuler.FlatStyle = FlatStyle.Flat
        AddHandler btnAnnuler.Click, AddressOf BtnAnnuler_Click
        Me.Controls.Add(btnAnnuler)

        Me.AcceptButton = btnOK
        Me.CancelButton = btnAnnuler
    End Sub

    Private Sub ChargerPreference()
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_KEY)
                If key Is Nothing Then Return
                Dim memorise As String = key.GetValue(REG_MEMORISE, "0").ToString()
                If memorise <> "1" Then Return
                cbMemoriser.Checked = True
                Dim fmt As String = key.GetValue(REG_FORMAT, "0").ToString()
                Select Case fmt
                    Case "1" : rbPdf.Checked = True
                    Case "2" : rbLesDeux.Checked = True
                    Case Else : rbDocx.Checked = True
                End Select
            End Using
        Catch
        End Try
    End Sub

    Private Sub SauvegarderPreference()
        Try
            Using key As RegistryKey = Registry.CurrentUser.CreateSubKey(REG_KEY)
                If cbMemoriser.Checked Then
                    Dim fmt As Integer = If(rbPdf.Checked, 1, If(rbLesDeux.Checked, 2, 0))
                    key.SetValue(REG_FORMAT, fmt.ToString())
                    key.SetValue(REG_MEMORISE, "1")
                Else
                    key.SetValue(REG_MEMORISE, "0")
                End If
            End Using
        Catch
        End Try
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        If rbPdf.Checked Then
            FormatChoisi = FormatSortie.PdfUniquement
        ElseIf rbLesDeux.Checked Then
            FormatChoisi = FormatSortie.DocxEtPdf
        Else
            FormatChoisi = FormatSortie.DocxUniquement
        End If
        SauvegarderPreference()
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub BtnAnnuler_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class
