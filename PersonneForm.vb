' ════════════════════════════════════════════════════════════════
' FichePersonnePhysiqueForm.vb
' Formulaire de création/modification d'une personne physique
' Validations : IPI 11 chiffres, COAD 6-9 chiffres, Tél +XXXXX,
'               N°Sécu 15 chiffres, Rôles multiples, TRIM partout
' ════════════════════════════════════════════════════════════════
Imports System.IO
Imports System.Windows.Forms
Imports System.Drawing
Imports OfficeOpenXml

Public Class PersonneForm
    Inherits Form

    ' ── Résultats publics ──────────────────────────────────────
    Public ReadOnly Roles         As List(Of String)   ' A, C, AR, AD
    Public ReadOnly Pseudonyme    As String
    Public ReadOnly Nom           As String
    Public ReadOnly Prenom        As String
    Public ReadOnly Genre         As String
    Public ReadOnly SocieteGestion As String
    Public ReadOnly IPI           As String
    Public ReadOnly IPI2          As String
    Public ReadOnly COAD          As String
    Public ReadOnly NumVoie       As String
    Public ReadOnly TypeVoie      As String
    Public ReadOnly NomVoie       As String
    Public ReadOnly CP            As String
    Public ReadOnly Ville         As String
    Public ReadOnly Pays          As String
    Public ReadOnly Mail          As String
    Public ReadOnly Tel           As String
    Public ReadOnly DateNaissance As String
    Public ReadOnly LieuNaissance As String
    Public ReadOnly NumSecu       As String

    ' ── Contexte grille (pour validation AR/AD) ───────────────
    Private ReadOnly _aDejaAuteur      As Boolean
    Private ReadOnly _aDejaCompositeur As Boolean

    ' ── Contrôles ─────────────────────────────────────────────
    Private pnlMain        As Panel
    Private lblTitre       As Label
    Private lblErreur      As Label

    ' Rôles
    Private cbAuteur       As CheckBox
    Private cbCompositeur  As CheckBox
    Private cbArrangeur    As CheckBox
    Private cbAdaptateur   As CheckBox

    ' Identité
    Private txtPseudo      As TextBox
    Private txtNom         As TextBox
    Private txtPrenom      As TextBox
    Private cboGenre       As ComboBox

    ' Société
    Private cboSociete     As ComboBox

    ' Numéros
    Private txtIPI         As TextBox
    Private txtIPI2        As TextBox
    Private txtCOAD        As TextBox

    ' Adresse
    Private txtNumVoie     As TextBox
    Private txtTypeVoie    As TextBox
    Private txtNomVoie     As TextBox
    Private txtCP          As TextBox
    Private txtVille       As TextBox
    Private txtPays        As TextBox

    ' Contact
    Private txtMail        As TextBox
    Private txtTel         As TextBox

    ' État civil
    Private txtDateNaiss   As TextBox
    Private txtLieuNaiss   As TextBox
    Private txtNumSecu     As TextBox

    ' Éditeurs
    Private txtEditeur     As TextBox
    Private _editeurValeur As String = ""   ' valeur brute XLSX "M00001:50;M00002:30"

    ' Boutons
    Private btnOK          As Button
    Private btnAnnuler     As Button

    ' ── Sociétés de gestion ───────────────────────────────────
    Private Shared ReadOnly Societes As String() = {
        "SACEM - France",
        "SAMRO - Afrique du Sud", "CAPASSO - Afrique du Sud",
        "ALBAUTOR - Albanie", "ONDA - Algérie", "GEMA - Allemagne",
        "UNAC - Angola", "SADAIC - Argentine", "ARMAUTHOR - Arménie",
        "AMCOS - Australie", "APRA - Australie",
        "AUSTRO MECHANA - Autriche", "AKM - Autriche",
        "AAS - Azerbaïdjan", "COSCAP - Barbade", "SABAM - Belgique",
        "BUBEDRA - Bénin", "SOBODAYCOM - Bolivie", "AMUS - Bosnie-Herzégovine",
        "UBC - Brésil", "SBAT - Brésil", "ABRAMUS - Brésil",
        "SADEMBRA - Brésil", "SBACEM - Brésil", "SOCINPRO - Brésil",
        "AMAR - Brésil", "ASSIM - Brésil",
        "MUSICAUTOR - Bulgarie", "BBDA - Burkina Faso",
        "SODRAC - Canada", "CMRRA - Canada", "SOCAN - Canada",
        "SCM - Cap-Vert", "SCD - Chili", "MCSC - Chine",
        "SAYCO - Colombie", "BCDA - Congo", "KOMCA - Corée du Sud",
        "ACAM - Costa Rica", "BURIDA - Côte d'Ivoire", "HDS ZAMP - Croatie",
        "ACDAM - Cuba", "KODA - Danemark", "NCB - Danemark Scandinavie",
        "SACERAU - Egypte", "SAYCE - Equateur",
        "SGAE - Espagne", "UNISON - Espagne", "SEDA - Espagne",
        "EAU - Estonie",
        "HFA - Etats-Unis", "SESAC - Etats-Unis", "AMRA - Etats-Unis",
        "ASCAP - Etats-Unis", "BMI - Etats-Unis",
        "TEOSTO - Finlande", "GCA - Géorgie",
        "EDEM - Grèce", "AUTODIA - Grèce", "ORFIUM - Grèce",
        "AEI - Guatemala", "BGDA - Guinée", "CASH - Hong Kong",
        "ARTISJUS - Hongrie", "MASA - Ile Maurice", "IPRS - Inde",
        "WAMI - Indonésie", "IMRO - Irlande", "STEF - Islande",
        "ACUM - Israël", "SIAE - Italie", "JACAP - Jamaïque",
        "JASRAC - Japon", "NexTone - Japon",
        "KAZAK - Kazakhstan", "AKKA / LAA - Lettonie", "LATGA-A - Lituanie",
        "MACA - Macao", "SOCOM / ZAMP - Macédoine", "OMDA - Madagascar",
        "MACP - Malaisie", "COSOMA - Malawi", "BUMDA - Mali",
        "BMDA - Maroc", "SACM - Mexique", "PAM CG - Monténégro",
        "NASCAM - Namibie", "MRCSN - Népal", "BNDA - Niger",
        "TONO - Norvège", "SACENC - Nouvelle-Calédonie",
        "SPAC - Panama", "APA - Paraguay",
        "STEMRA - Pays-Bas", "BUMA - Pays-Bas",
        "APDAYC - Pérou", "FILSCAP - Philippines", "ZAIKS - Pologne",
        "SPA - Portugal", "BUCADA - République Centrafricaine",
        "OSA - République Tchèque", "UCMR-ADA - Roumanie",
        "MCPS - Royaume Uni", "PRS - Royaume Uni",
        "RAO - Russie", "ECCO - Sainte-Lucie", "SODAV - Sénégal",
        "SOKOJ - Serbie", "COMPASS - Singapour", "SOZA - Slovaquie",
        "SAZAS - Slovénie", "STIM - Suède",
        "SUISA - Suisse", "PRO-LITTERIS - Suisse", "SSA - Suisse",
        "MÜST - Taiwan", "MCT - Thaïlande", "BUTODRA - Togo",
        "COTT - Trinité et Tobago", "MESAM - Turquie", "MSG - Turquie",
        "UACRR - Ukraine", "AGADU - Uruguay", "SACVEN - Venezuela",
        "VCPMC - Vietnam", "ZAMCOPS - Zambie"
    }

    ' ── Id pré-calculé (optionnel, affiché dans le titre) ───────
    Public Property ChampsId As String = ""

    ' ── Constructeur ──────────────────────────────────────────
    Public Sub New(aDejaAuteur As Boolean, aDejaCompositeur As Boolean,
                   Optional existingRow As DataRow = Nothing)
        _aDejaAuteur      = aDejaAuteur
        _aDejaCompositeur = aDejaCompositeur
        Roles = New List(Of String)()
        InitializeComponent()
        If existingRow IsNot Nothing Then ChargerExistant(existingRow)
    End Sub

    ' ── UI ────────────────────────────────────────────────────
    Private Sub InitializeComponent()
        Me.Text            = "Fiche Personne Physique"
        Me.Size            = New Size(620, 780)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.BackColor       = Color.White

        ' Scroll panel
        pnlMain = New Panel() With {
            .AutoScroll = True,
            .Dock       = DockStyle.Fill,
            .Padding    = New Padding(15)
        }
        Me.Controls.Add(pnlMain)

        Dim y As Integer = 10

        ' ── Titre ──
        lblTitre = New Label() With {
            .Text      = "NOUVELLE PERSONNE PHYSIQUE",
            .Font      = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.FromArgb(60, 0, 100),
            .Location  = New Point(10, y),
            .Size      = New Size(570, 28)
        }
        pnlMain.Controls.Add(lblTitre)
        y += 38

        ' ── Erreur ──
        lblErreur = New Label() With {
            .Text      = "",
            .ForeColor = Color.DarkRed,
            .Font      = New Font("Segoe UI", 8.5, FontStyle.Bold),
            .Location  = New Point(10, y),
            .Size      = New Size(570, 40),
            .Visible   = False
        }
        pnlMain.Controls.Add(lblErreur)
        y += 45

        ' ── Section Rôles ──
        y = AddSectionHeader("RÔLES", y)
        Dim pnlRoles As New Panel() With {
            .Location  = New Point(10, y),
            .Size      = New Size(570, 36),
            .BackColor = Color.FromArgb(245, 240, 255)
        }
        cbAuteur      = NewCheckBox("Auteur",      pnlRoles, 5)
        cbCompositeur = NewCheckBox("Compositeur", pnlRoles, 145)
        cbArrangeur   = NewCheckBox("Arrangeur",   pnlRoles, 285)
        cbAdaptateur  = NewCheckBox("Adaptateur",  pnlRoles, 415)
        pnlMain.Controls.Add(pnlRoles)
        y += 46

        ' ── Section Identité ──
        y = AddSectionHeader("IDENTITÉ", y)
        txtPseudo = AddField("Pseudonyme", y, pnlMain) : y += 30
        txtNom    = AddField("Nom *", y, pnlMain)      : y += 30
        txtPrenom = AddField("Prénom *", y, pnlMain)   : y += 30
        cboGenre  = AddCombo("Genre", y, pnlMain, {"Masculin", "Féminin", "Autre"}) : y += 30

        ' ── Section Société ──
        y = AddSectionHeader("SOCIÉTÉ DE GESTION", y)
        cboSociete = AddCombo("Société *", y, pnlMain, Societes) : y += 30

        ' ── Section Numéros ──
        y = AddSectionHeader("NUMÉROS", y)

        Dim lblIPIHint As New Label() With {
            .Text     = "IPI : 11 chiffres (ex: 00123456789)   COAD : 6 à 9 chiffres",
            .Font     = New Font("Segoe UI", 7.5),
            .ForeColor = Color.Gray,
            .Location = New Point(10, y),
            .Size     = New Size(570, 16)
        }
        pnlMain.Controls.Add(lblIPIHint)
        y += 18

        txtIPI  = AddField("IPI",   y, pnlMain) : y += 30
        txtIPI2 = AddField("IPI 2", y, pnlMain) : y += 30
        txtCOAD = AddField("COAD",  y, pnlMain) : y += 30

        ' ── Section Éditeurs ──
        y = AddSectionHeader("ÉDITEURS ASSOCIÉS", y)
        Dim lblEditeur As New Label() With {
            .Text     = "Éditeur(s)",
            .Location = New Point(10, y + 4),
            .Size     = New Size(160, 20),
            .Font     = New Font("Segoe UI", 8.5)
        }
        txtEditeur = New TextBox() With {
            .Location  = New Point(175, y),
            .Size      = New Size(330, 24),
            .Font      = New Font("Segoe UI", 9),
            .ReadOnly  = True,
            .BackColor = Color.White
        }
        Dim btnEditeur As New Button() With {
            .Text      = "...",
            .Location  = New Point(510, y),
            .Size      = New Size(50, 24),
            .FlatStyle = FlatStyle.Flat
        }
        pnlMain.Controls.Add(lblEditeur)
        pnlMain.Controls.Add(txtEditeur)
        pnlMain.Controls.Add(btnEditeur)
        AddHandler btnEditeur.Click, Sub(s, ev)
                                         Using dlg As New EditeurListForm(_editeurValeur)
                                             If dlg.ShowDialog() = DialogResult.OK Then
                                                 _editeurValeur  = dlg.ValeurEditeurs
                                                 txtEditeur.Text = dlg.NomEditeurs
                                             End If
                                         End Using
                                     End Sub
        y += 30

        ' ── Section Adresse ──
        y = AddSectionHeader("ADRESSE", y)
        txtNumVoie  = AddField("N° voie",   y, pnlMain) : y += 30
        txtTypeVoie = AddField("Type voie", y, pnlMain) : y += 30
        txtNomVoie  = AddField("Nom voie",  y, pnlMain) : y += 30
        txtCP       = AddField("Code postal", y, pnlMain) : y += 30
        txtVille    = AddField("Ville",     y, pnlMain) : y += 30
        txtPays     = AddField("Pays",      y, pnlMain) : y += 30

        ' ── Section Contact ──
        y = AddSectionHeader("CONTACT", y)

        Dim lblTelHint As New Label() With {
            .Text      = "Téléphone : format international +33612345678",
            .Font      = New Font("Segoe UI", 7.5),
            .ForeColor = Color.Gray,
            .Location  = New Point(10, y),
            .Size      = New Size(570, 16)
        }
        pnlMain.Controls.Add(lblTelHint)
        y += 18

        txtMail = AddField("E-mail", y, pnlMain) : y += 30
        txtTel  = AddField("Téléphone", y, pnlMain) : y += 30

        ' ── Section État civil ──
        y = AddSectionHeader("ÉTAT CIVIL", y)

        Dim lblSecuHint As New Label() With {
            .Text      = "N° Sécurité sociale : 15 chiffres sans espace",
            .Font      = New Font("Segoe UI", 7.5),
            .ForeColor = Color.Gray,
            .Location  = New Point(10, y),
            .Size      = New Size(570, 16)
        }
        pnlMain.Controls.Add(lblSecuHint)
        y += 18

        txtDateNaiss = AddField("Date naissance (JJ/MM/AAAA)", y, pnlMain) : y += 30
        txtLieuNaiss = AddField("Lieu naissance", y, pnlMain)              : y += 30
        txtNumSecu   = AddField("N° Sécurité sociale", y, pnlMain)         : y += 30

        ' ── Boutons ──
        y += 10
        btnOK = New Button() With {
            .Text      = "OK",
            .Location  = New Point(390, y),
            .Size      = New Size(90, 32),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font      = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnAnnuler = New Button() With {
            .Text      = "Annuler",
            .Location  = New Point(490, y),
            .Size      = New Size(90, 32),
            .FlatStyle = FlatStyle.Flat
        }
        pnlMain.Controls.Add(btnOK)
        pnlMain.Controls.Add(btnAnnuler)

        ' Hauteur panel intérieur
        pnlMain.AutoScrollMinSize = New Size(580, y + 50)

        AddHandler btnOK.Click,      AddressOf BtnOK_Click
        AddHandler btnAnnuler.Click, Sub(s, e) Me.DialogResult = DialogResult.Cancel
        Me.AcceptButton = btnOK
        Me.CancelButton = btnAnnuler
    End Sub

    ' ── Helpers UI ────────────────────────────────────────────
    Private Function AddSectionHeader(title As String, y As Integer) As Integer
        Dim lbl As New Label() With {
            .Text      = title,
            .Font      = New Font("Segoe UI", 8, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(60, 0, 100),
            .Location  = New Point(10, y),
            .Size      = New Size(570, 20),
            .Padding   = New Padding(4, 2, 0, 0)
        }
        pnlMain.Controls.Add(lbl)
        Return y + 24
    End Function

    Private Function AddField(label As String, y As Integer, parent As Control) As TextBox
        Dim lbl As New Label() With {
            .Text     = label,
            .Location = New Point(10, y + 4),
            .Size     = New Size(160, 20),
            .Font     = New Font("Segoe UI", 8.5)
        }
        Dim txt As New TextBox() With {
            .Location = New Point(175, y),
            .Size     = New Size(395, 24),
            .Font     = New Font("Segoe UI", 9)
        }
        parent.Controls.Add(lbl)
        parent.Controls.Add(txt)
        Return txt
    End Function

    Private Function AddCombo(label As String, y As Integer, parent As Control,
                               items As String()) As ComboBox
        Dim lbl As New Label() With {
            .Text     = label,
            .Location = New Point(10, y + 4),
            .Size     = New Size(160, 20),
            .Font     = New Font("Segoe UI", 8.5)
        }
        Dim cbo As New ComboBox() With {
            .Location     = New Point(175, y),
            .Size         = New Size(395, 24),
            .Font         = New Font("Segoe UI", 9),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cbo.Items.AddRange(items)
        parent.Controls.Add(lbl)
        parent.Controls.Add(cbo)
        Return cbo
    End Function

    Private Function NewCheckBox(text As String, parent As Control, x As Integer) As CheckBox
        Dim cb As New CheckBox() With {
            .Text     = text,
            .Location = New Point(x, 8),
            .Size     = New Size(130, 20),
            .Font     = New Font("Segoe UI", 8.5)
        }
        parent.Controls.Add(cb)
        Return cb
    End Function

    ' ── Chargement existant ───────────────────────────────────
    Private Sub ChargerExistant(row As DataRow)
        lblTitre.Text = "MODIFIER PERSONNE PHYSIQUE"

        ' Rôles
        Dim roles As String = SafeStr(row, "Role").Trim().ToUpper()
        cbAuteur.Checked      = roles.Contains("A") AndAlso Not roles.Contains("AR") AndAlso Not roles.Contains("AD")
        cbCompositeur.Checked = roles.Contains("C")
        cbArrangeur.Checked   = roles.Contains("AR")
        cbAdaptateur.Checked  = roles.Contains("AD")

        txtPseudo.Text    = SafeStr(row, "Pseudonyme")
        txtNom.Text       = SafeStr(row, "Nom")
        txtPrenom.Text    = SafeStr(row, "Prenom")
        cboGenre.Text     = SafeStr(row, "Genre")

        ' Société — chercher par code
        Dim soc As String = SafeStr(row, "SocieteGestion").Trim().ToUpper()
        For Each item As String In Societes
            If item.ToUpper().StartsWith(soc) Then
                cboSociete.SelectedItem = item
                Exit For
            End If
        Next

        _editeurValeur    = SafeStr(row, "Editeur")
        txtEditeur.Text   = ResolveEditeurIdsToNames(_editeurValeur)
        txtIPI.Text       = SafeStr(row, "IPI")
        txtIPI2.Text      = SafeStr(row, "IPI 2")
        txtCOAD.Text      = SafeStr(row, "COAD")
        txtNumVoie.Text   = SafeStr(row, "Num de voie")
        txtTypeVoie.Text  = SafeStr(row, "Type de voie")
        txtNomVoie.Text   = SafeStr(row, "Nom de voie")
        txtCP.Text        = SafeStr(row, "CP")
        txtVille.Text     = SafeStr(row, "Ville")
        txtPays.Text      = SafeStr(row, "Pays")
        txtMail.Text      = SafeStr(row, "Mail")
        txtTel.Text       = SafeStr(row, "Tel")
        txtDateNaiss.Text = SafeStr(row, "Date de naissance")
        txtLieuNaiss.Text = SafeStr(row, "Lieu de naissance")
        txtNumSecu.Text   = SafeStr(row, "Num de Securite Sociale")
    End Sub

    Private Function SafeStr(row As DataRow, col As String) As String
        Try : Return row(col).ToString() : Catch : Return "" : End Try
    End Function

    ' ── Validation et OK ─────────────────────────────────────
    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        lblErreur.Visible = False
        Dim erreurs As New List(Of String)()

        ' ── Rôles ──
        Dim rolesSelectionnes As New List(Of String)()
        If cbAuteur.Checked      Then rolesSelectionnes.Add("A")
        If cbCompositeur.Checked Then rolesSelectionnes.Add("C")
        If cbArrangeur.Checked   Then rolesSelectionnes.Add("AR")
        If cbAdaptateur.Checked  Then rolesSelectionnes.Add("AD")

        If rolesSelectionnes.Count = 0 Then
            erreurs.Add("• Sélectionnez au moins un rôle.")
        End If
        If cbArrangeur.Checked AndAlso Not _aDejaCompositeur AndAlso Not cbCompositeur.Checked Then
            erreurs.Add("• Arrangeur requiert un Compositeur dans la grille.")
        End If
        If cbAdaptateur.Checked AndAlso Not _aDejaAuteur AndAlso Not cbAuteur.Checked Then
            erreurs.Add("• Adaptateur requiert un Auteur dans la grille.")
        End If

        ' ── Nom/Prénom ──
        Dim nom    As String = txtNom.Text.Trim()
        Dim prenom As String = txtPrenom.Text.Trim()
        If String.IsNullOrEmpty(nom) OrElse String.IsNullOrEmpty(prenom) Then
            erreurs.Add("• Nom et Prénom sont obligatoires.")
        End If

        ' ── Société ──
        If cboSociete.SelectedIndex < 0 Then
            erreurs.Add("• Sélectionnez une société de gestion.")
        End If

        ' ── IPI : 11 chiffres, padding auto ──
        Dim ipiVal As String = txtIPI.Text.Trim().Replace(" ", "")
        If Not String.IsNullOrEmpty(ipiVal) Then
            ipiVal = ipiVal.PadLeft(11, "0"c)
            If ipiVal.Length <> 11 OrElse Not ipiVal.All(Function(c) Char.IsDigit(c)) Then
                erreurs.Add("• IPI invalide (11 chiffres max).")
            End If
        End If

        Dim ipi2Val As String = txtIPI2.Text.Trim().Replace(" ", "")
        If Not String.IsNullOrEmpty(ipi2Val) Then
            ipi2Val = ipi2Val.PadLeft(11, "0"c)
            If ipi2Val.Length <> 11 OrElse Not ipi2Val.All(Function(c) Char.IsDigit(c)) Then
                erreurs.Add("• IPI 2 invalide (11 chiffres max).")
            End If
        End If

        ' ── COAD : 6 à 9 chiffres ──
        Dim coadVal As String = txtCOAD.Text.Trim().Replace(" ", "").ToUpper()
        If Not String.IsNullOrEmpty(coadVal) Then
            If Not coadVal.All(Function(c) Char.IsDigit(c)) OrElse
               coadVal.Length < 6 OrElse coadVal.Length > 9 Then
                erreurs.Add("• COAD invalide (6 à 9 chiffres).")
            End If
        End If

        ' ── Téléphone : +XXXXX ──
        Dim telVal As String = txtTel.Text.Trim().Replace(" ", "")
        If Not String.IsNullOrEmpty(telVal) Then
            If Not telVal.StartsWith("+") OrElse
               Not telVal.Substring(1).All(Function(c) Char.IsDigit(c)) OrElse
               telVal.Length < 8 Then
                erreurs.Add("• Téléphone invalide (format +33612345678).")
            End If
        End If

        ' ── N° Sécu : 15 chiffres ──
        Dim secuVal As String = txtNumSecu.Text.Trim().Replace(" ", "")
        If Not String.IsNullOrEmpty(secuVal) Then
            If Not secuVal.All(Function(c) Char.IsDigit(c)) OrElse secuVal.Length <> 15 Then
                erreurs.Add("• N° Sécu invalide (exactement 15 chiffres).")
            End If
        End If

        ' ── Afficher erreurs ou valider ──
        If erreurs.Count > 0 Then
            lblErreur.Text    = String.Join(Environment.NewLine, erreurs)
            lblErreur.Visible = True
            Return
        End If

        ' ── Affecter résultats (avec TRIM partout) ──
        Roles.AddRange(rolesSelectionnes)

        ' Extraire code société (ex: "SACEM - France" → "SACEM")
        Dim socRaw As String = If(cboSociete.SelectedItem?.ToString(), "")
        Dim socCode As String = If(socRaw.Contains(" - "), socRaw.Split({" - "}, StringSplitOptions.None)(0).Trim(), socRaw.Trim())

        ' Propriétés via reflection impossible en VB → on passe par dictionnaire
        ' → Les valeurs sont lues via les propriétés publiques ReadOnly
        ' On utilise SetProprietes()
        SetProprietes(nom, prenom, socCode, ipiVal, ipi2Val, coadVal, telVal, secuVal)

        Me.DialogResult = DialogResult.OK
    End Sub

    ' ── Setters (contournement ReadOnly via Friend) ───────────
    Private _pseudo    As String : Private _nom       As String
    Private _prenom    As String : Private _genre     As String
    Private _societe   As String : Private _ipi       As String
    Private _ipi2      As String : Private _coad      As String
    Private _numVoie   As String : Private _typeVoie  As String
    Private _nomVoie   As String : Private _cp        As String
    Private _ville     As String : Private _pays      As String
    Private _mail      As String : Private _tel       As String
    Private _dateNaiss As String : Private _lieuNaiss As String
    Private _numSecu   As String


    Private Function ResolveEditeurIdsToNames(raw As String) As String
        If String.IsNullOrEmpty(raw) Then Return ""
        Dim dtMor As DataTable = Nothing
        Try
            If File.Exists(PersonnesForm.ResolveXlsxPath()) Then
                Using pkg As New ExcelPackage(New FileInfo(PersonnesForm.ResolveXlsxPath()))
                    Dim ws = pkg.Workbook.Worksheets("PERSONNEMORALE")
                    If ws IsNot Nothing AndAlso ws.Dimension IsNot Nothing Then
                        dtMor = New DataTable()
                        For c = 1 To ws.Dimension.Columns
                            dtMor.Columns.Add(ws.Cells(1, c).Text.Trim())
                        Next
                        For r = 2 To ws.Dimension.Rows
                            Dim nr As DataRow = dtMor.NewRow()
                            For c = 1 To ws.Dimension.Columns
                                nr(c - 1) = ws.Cells(r, c).Text.Trim()
                            Next
                            dtMor.Rows.Add(nr)
                        Next
                    End If
                End Using
            End If
        Catch
        End Try
        Dim parts As New List(Of String)()
        For Each entry As String In raw.Split(";"c)
            Dim t As String = entry.Trim()
            If String.IsNullOrEmpty(t) Then Continue For
            Dim colonIdx As Integer = t.LastIndexOf(":"c)
            Dim idOuNom As String = t
            Dim pct As String = ""
            If colonIdx > 0 Then
                Dim droite As String = t.Substring(colonIdx + 1).Trim()
                Dim testPct As Double
                If Double.TryParse(droite, Globalization.NumberStyles.Any,
                                   Globalization.CultureInfo.InvariantCulture, testPct) Then
                    idOuNom = t.Substring(0, colonIdx).Trim()
                    pct = droite
                End If
            End If
            Dim displayNom As String = idOuNom
            If dtMor IsNot Nothing AndAlso dtMor.Columns.Contains("Id") AndAlso dtMor.Columns.Contains("Designation") Then
                For Each mr As DataRow In dtMor.Rows
                    If mr("Id").ToString().Trim().ToUpper() = idOuNom.ToUpper() Then
                        displayNom = mr("Designation").ToString().Trim()
                        Exit For
                    End If
                Next
            End If
            parts.Add(If(String.IsNullOrEmpty(pct), displayNom, displayNom & " (" & pct & "%)"))
        Next
        Return String.Join(" ; ", parts)
    End Function
    Private Sub SetProprietes(nom As String, prenom As String, societe As String,
                               ipi As String, ipi2 As String, coad As String,
                               tel As String, secu As String)
        _nom       = nom
        ' _editeurValeur déjà mis à jour via btnEditeur.Click
        _prenom    = prenom
        _pseudo    = txtPseudo.Text.Trim()
        _genre     = If(cboGenre.SelectedItem?.ToString(), "").Trim()
        _societe   = societe
        _ipi       = ipi
        _ipi2      = ipi2
        _coad      = coad
        _numVoie   = txtNumVoie.Text.Trim()
        _typeVoie  = txtTypeVoie.Text.Trim()
        _nomVoie   = txtNomVoie.Text.Trim()
        _cp        = txtCP.Text.Trim()
        _ville     = txtVille.Text.Trim()
        _pays      = txtPays.Text.Trim()
        _mail      = txtMail.Text.Trim()
        _tel       = tel
        _dateNaiss = txtDateNaiss.Text.Trim()
        _lieuNaiss = txtLieuNaiss.Text.Trim()
        _numSecu   = secu
    End Sub

    ' ── Propriétés publiques (lecture résultats) ──────────────
    Public ReadOnly Property ResultEditeur() As String
        Get
            Return _editeurValeur
        End Get
    End Property
    Public ReadOnly Property ResultNom() As String
        Get
            Return _nom
        End Get
    End Property
    Public ReadOnly Property ResultPrenom() As String
        Get
            Return _prenom
        End Get
    End Property
    Public ReadOnly Property ResultPseudo() As String
        Get
            Return _pseudo
        End Get
    End Property
    Public ReadOnly Property ResultGenre() As String
        Get
            Return _genre
        End Get
    End Property
    Public ReadOnly Property ResultSociete() As String
        Get
            Return _societe
        End Get
    End Property
    Public ReadOnly Property ResultIPI() As String
        Get
            Return _ipi
        End Get
    End Property
    Public ReadOnly Property ResultIPI2() As String
        Get
            Return _ipi2
        End Get
    End Property
    Public ReadOnly Property ResultCOAD() As String
        Get
            Return _coad
        End Get
    End Property
    Public ReadOnly Property ResultNumVoie() As String
        Get
            Return _numVoie
        End Get
    End Property
    Public ReadOnly Property ResultTypeVoie() As String
        Get
            Return _typeVoie
        End Get
    End Property
    Public ReadOnly Property ResultNomVoie() As String
        Get
            Return _nomVoie
        End Get
    End Property
    Public ReadOnly Property ResultCP() As String
        Get
            Return _cp
        End Get
    End Property
    Public ReadOnly Property ResultVille() As String
        Get
            Return _ville
        End Get
    End Property
    Public ReadOnly Property ResultPays() As String
        Get
            Return _pays
        End Get
    End Property
    Public ReadOnly Property ResultMail() As String
        Get
            Return _mail
        End Get
    End Property
    Public ReadOnly Property ResultTel() As String
        Get
            Return _tel
        End Get
    End Property
    Public ReadOnly Property ResultDateNaiss() As String
        Get
            Return _dateNaiss
        End Get
    End Property
    Public ReadOnly Property ResultLieuNaiss() As String
        Get
            Return _lieuNaiss
        End Get
    End Property
    Public ReadOnly Property ResultNumSecu() As String
        Get
            Return _numSecu
        End Get
    End Property

End Class
