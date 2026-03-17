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

    ' IPI (DGV)
    Private dgvIPI         As DataGridView
    Private _dtIPI         As DataTable
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
        Me.Size            = New Size(960, 590)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.BackColor       = Color.White

        pnlMain = New Panel() With {
            .AutoScroll = False,
            .Dock       = DockStyle.Fill
        }
        Me.Controls.Add(pnlMain)

        ' ════════════════════════════════════════
        ' COLONNE GAUCHE  x=10  w=560
        ' ════════════════════════════════════════
        Dim L As Integer = 10
        Dim LW As Integer = 560
        Dim y As Integer = 10

        ' Titre
        lblTitre = New Label() With {
            .Text      = "NOUVELLE PERSONNE PHYSIQUE",
            .Font      = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.FromArgb(60, 0, 100),
            .Location  = New Point(L, y),
            .Size      = New Size(LW, 26)
        }
        pnlMain.Controls.Add(lblTitre)
        y += 30

        ' Erreur
        lblErreur = New Label() With {
            .Text      = "",
            .ForeColor = Color.DarkRed,
            .Font      = New Font("Segoe UI", 8.5, FontStyle.Bold),
            .Location  = New Point(L, y),
            .Size      = New Size(LW, 34),
            .Visible   = False
        }
        pnlMain.Controls.Add(lblErreur)
        y += 36

        ' Roles
        y = AddSectionHeader(y, L, LW, "RÔLES")
        Dim pnlRoles As New Panel() With {
            .Location  = New Point(L, y),
            .Size      = New Size(LW, 28),
            .BackColor = Color.FromArgb(245, 240, 255)
        }
        cbAuteur      = NewCheckBox("Auteur",      pnlRoles, 5)
        cbCompositeur = NewCheckBox("Compositeur", pnlRoles, 130)
        cbArrangeur   = NewCheckBox("Arrangeur",   pnlRoles, 255)
        cbAdaptateur  = NewCheckBox("Adaptateur",  pnlRoles, 380)
        pnlMain.Controls.Add(pnlRoles)
        y += 36

        ' Identite
        y = AddSectionHeader(y, L, LW, "IDENTITÉ")
        txtPseudo = AddField2("Pseudonyme", y, L, LW) : y += 26
        txtNom    = AddField2("Nom *",      y, L, LW) : y += 26
        txtPrenom = AddField2("Prénom *",   y, L, LW) : y += 26
        cboGenre  = AddCombo2("Genre",      y, L, LW, {"MR", "MME", "Autre"}) : y += 26

        ' Societe
        y = AddSectionHeader(y, L, LW, "SOCIÉTÉ DE GESTION")
        cboSociete = AddCombo2("Société *", y, L, LW, Societes) : y += 26

        ' Numeros
        y = AddSectionHeader(y, L, LW, "NUMÉROS")
        pnlMain.Controls.Add(New Label() With {
            .Text      = "IPI : 11 chiffres max   COAD : 6 à 9 chiffres",
            .Font      = New Font("Segoe UI", 7.5),
            .ForeColor = Color.Gray,
            .Location  = New Point(L, y),
            .Size      = New Size(LW, 13)
        })
        y += 14

        pnlMain.Controls.Add(New Label() With {
            .Text     = "IPI(s)",
            .Location = New Point(L, y + 2),
            .Size     = New Size(120, 18),
            .Font     = New Font("Segoe UI", 8.5)
        })

        _dtIPI = New DataTable()
        _dtIPI.Columns.Add("Rôle(s)", GetType(String))
        _dtIPI.Columns.Add("IPI",     GetType(String))
        _dtIPI.Columns.Add("Nom",     GetType(String))

        dgvIPI = New DataGridView() With {
            .Location              = New Point(L + 120, y),
            .Size                  = New Size(350, 78),
            .DataSource            = _dtIPI,
            .AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible     = False,
            .AllowUserToAddRows    = False,
            .AllowUserToDeleteRows = False,
            .Font                  = New Font("Segoe UI", 8F),
            .BorderStyle           = BorderStyle.FixedSingle,
            .BackgroundColor       = Color.White,
            .SelectionMode         = DataGridViewSelectionMode.FullRowSelect
        }
        pnlMain.Controls.Add(dgvIPI)

        Dim btnIPIAdd As New Button() With {
            .Text      = "+",
            .Location  = New Point(L + 474, y),
            .Size      = New Size(24, 24),
            .BackColor = Color.FromArgb(0, 120, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font      = New Font("Segoe UI", 8, FontStyle.Bold)
        }
        Dim btnIPIDel As New Button() With {
            .Text      = "−",
            .Location  = New Point(L + 474, y + 27),
            .Size      = New Size(24, 24),
            .BackColor = Color.FromArgb(180, 30, 30),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font      = New Font("Segoe UI", 8, FontStyle.Bold)
        }
        Dim btnIPISACEM As New Button() With {
            .Text      = "🔍",
            .Location  = New Point(L + 474, y + 54),
            .Size      = New Size(24, 24),
            .BackColor = Color.FromArgb(20, 60, 140),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font      = New Font("Segoe UI", 8)
        }
        pnlMain.Controls.Add(btnIPIAdd)
        pnlMain.Controls.Add(btnIPIDel)
        pnlMain.Controls.Add(btnIPISACEM)

        AddHandler btnIPIAdd.Click, Sub(s, ev)
            Dim nr As DataRow = _dtIPI.NewRow()
            nr("Rôle(s)") = "" : nr("IPI") = "" : nr("Nom") = ""
            _dtIPI.Rows.Add(nr)
        End Sub
        AddHandler btnIPIDel.Click, Sub(s, ev)
            If dgvIPI.SelectedRows.Count > 0 Then
                Dim idx As Integer = dgvIPI.SelectedRows(0).Index
                If idx >= 0 AndAlso idx < _dtIPI.Rows.Count Then _dtIPI.Rows.RemoveAt(idx)
            End If
        End Sub
        AddHandler btnIPISACEM.Click, AddressOf BtnIPISACEM_Click

        y += 82
        txtCOAD = AddField2("COAD", y, L, LW) : y += 26

        ' Editeurs
        y = AddSectionHeader(y, L, LW, "ÉDITEURS ASSOCIÉS")
        pnlMain.Controls.Add(New Label() With {
            .Text     = "Éditeur(s)",
            .Location = New Point(L, y + 3),
            .Size     = New Size(120, 18),
            .Font     = New Font("Segoe UI", 8.5)
        })
        txtEditeur = New TextBox() With {
            .Location  = New Point(L + 120, y),
            .Size      = New Size(370, 22),
            .Font      = New Font("Segoe UI", 9),
            .ReadOnly  = True,
            .BackColor = Color.White
        }
        Dim btnEditeur As New Button() With {
            .Text      = "...",
            .Location  = New Point(L + 494, y),
            .Size      = New Size(36, 22),
            .FlatStyle = FlatStyle.Flat
        }
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

        ' ════════════════════════════════════════
        ' COLONNE DROITE  x=590  w=340
        ' ════════════════════════════════════════
        Dim R As Integer = 590
        Dim RW As Integer = 340
        Dim yr As Integer = 10

        ' Adresse
        yr = AddSectionHeader(yr, R, RW, "ADRESSE")
        txtNumVoie  = AddField2("N° voie",     yr, R, RW) : yr += 26
        txtTypeVoie = AddField2("Type voie",   yr, R, RW) : yr += 26
        txtNomVoie  = AddField2("Nom voie",    yr, R, RW) : yr += 26
        txtCP       = AddField2("Code postal", yr, R, RW) : yr += 26
        txtVille    = AddField2("Ville",       yr, R, RW) : yr += 26
        txtPays     = AddField2("Pays",        yr, R, RW) : yr += 26

        ' Contact
        yr = AddSectionHeader(yr, R, RW, "CONTACT")
        pnlMain.Controls.Add(New Label() With {
            .Text      = "Format : +33612345678",
            .Font      = New Font("Segoe UI", 7.5),
            .ForeColor = Color.Gray,
            .Location  = New Point(R, yr),
            .Size      = New Size(RW, 13)
        })
        yr += 14
        txtMail = AddField2("E-mail",    yr, R, RW) : yr += 26
        txtTel  = AddField2("Téléphone", yr, R, RW) : yr += 26

        ' Etat civil
        yr = AddSectionHeader(yr, R, RW, "ÉTAT CIVIL")
        pnlMain.Controls.Add(New Label() With {
            .Text      = "N° Sécu : 15 chiffres sans espace",
            .Font      = New Font("Segoe UI", 7.5),
            .ForeColor = Color.Gray,
            .Location  = New Point(R, yr),
            .Size      = New Size(RW, 13)
        })
        yr += 14
        txtDateNaiss = AddField2("Date naissance (JJ/MM/AAAA)", yr, R, RW) : yr += 26
        txtLieuNaiss = AddField2("Lieu naissance",              yr, R, RW) : yr += 26
        txtNumSecu   = AddField2("N° Sécurité sociale",         yr, R, RW) : yr += 26

        ' Boutons OK / Annuler
        yr += 10
        btnOK = New Button() With {
            .Text      = "OK",
            .Location  = New Point(R + 120, yr),
            .Size      = New Size(90, 28),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font      = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnAnnuler = New Button() With {
            .Text      = "Annuler",
            .Location  = New Point(R + 218, yr),
            .Size      = New Size(90, 28),
            .FlatStyle = FlatStyle.Flat
        }
        pnlMain.Controls.Add(btnOK)
        pnlMain.Controls.Add(btnAnnuler)

        AddHandler btnOK.Click,      AddressOf BtnOK_Click
        AddHandler btnAnnuler.Click, Sub(s, e) Me.DialogResult = DialogResult.Cancel
        Me.AcceptButton = btnOK
        Me.CancelButton = btnAnnuler
    End Sub

    ' ── Helpers UI ────────────────────────────────────────────
    ' Helpers UI — version 2 colonnes (x, w paramétrables)
    Private Function AddSectionHeader(y As Integer, x As Integer, w As Integer, title As String) As Integer
        pnlMain.Controls.Add(New Label() With {
            .Text      = title,
            .Font      = New Font("Segoe UI", 8, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(60, 0, 100),
            .Location  = New Point(x, y),
            .Size      = New Size(w, 18),
            .Padding   = New Padding(3, 1, 0, 0)
        })
        Return y + 22
    End Function

    Private Function AddField2(label As String, y As Integer, x As Integer, w As Integer) As TextBox
        Dim lblW As Integer = CInt(w * 0.38)
        pnlMain.Controls.Add(New Label() With {
            .Text     = label,
            .Location = New Point(x, y + 3),
            .Size     = New Size(lblW, 18),
            .Font     = New Font("Segoe UI", 8.5)
        })
        Dim txt As New TextBox() With {
            .Location = New Point(x + lblW, y),
            .Size     = New Size(w - lblW, 22),
            .Font     = New Font("Segoe UI", 9)
        }
        pnlMain.Controls.Add(txt)
        Return txt
    End Function

    Private Function AddCombo2(label As String, y As Integer, x As Integer, w As Integer,
                                items As String()) As ComboBox
        Dim lblW As Integer = CInt(w * 0.38)
        pnlMain.Controls.Add(New Label() With {
            .Text     = label,
            .Location = New Point(x, y + 3),
            .Size     = New Size(lblW, 18),
            .Font     = New Font("Segoe UI", 8.5)
        })
        Dim cbo As New ComboBox() With {
            .Location      = New Point(x + lblW, y),
            .Size          = New Size(w - lblW, 22),
            .Font          = New Font("Segoe UI", 9),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cbo.Items.AddRange(items)
        pnlMain.Controls.Add(cbo)
        Return cbo
    End Function

    ' Anciens helpers conservés pour compatibilité (ChargerExistant, etc.)
    Private Function AddField(label As String, y As Integer, parent As Control) As TextBox
        Return AddField2(label, y, 10, 560)
    End Function
    Private Function AddCombo(label As String, y As Integer, parent As Control,
                               items As String()) As ComboBox
        Return AddCombo2(label, y, 10, 560, items)
    End Function

    ' ── Recherche IPI dans répertoire SACEM ───────────────────
    Private Sub BtnIPISACEM_Click(sender As Object, e As EventArgs)
        Dim pseudo As String = If(txtPseudo IsNot Nothing, txtPseudo.Text.Trim(), "")
        Dim nom    As String = If(txtNom    IsNot Nothing, txtNom.Text.Trim(),    "")
        Dim prenom As String = If(txtPrenom IsNot Nothing, txtPrenom.Text.Trim(), "")

        Dim query As String = String.Join(" ", {pseudo, nom, prenom}.Where(Function(s) Not String.IsNullOrEmpty(s)))
        If String.IsNullOrEmpty(query) Then
            MessageBox.Show("Saisir au moins Nom ou Prénom.", "Recherche SACEM", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Trouver le script Python
        Dim scriptPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "Scripts", "sacem_repertoire_public.py")
        scriptPath = IO.Path.GetFullPath(scriptPath)
        If Not IO.File.Exists(scriptPath) Then
            MessageBox.Show("Script introuvable : " & scriptPath, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim results As New List(Of Tuple(Of String, String, String))() ' IPI, Nom, Roles
        Dim errMsg As String = ""

        Using dlgWait As New Form() With {
            .Text = "Recherche SACEM…",
            .Size = New Size(340, 80),
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedToolWindow
        }
            Dim lbl As New Label() With {
                .Text = "Recherche en cours : " & query,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter
            }
            dlgWait.Controls.Add(lbl)
            dlgWait.Show(Me)
            Application.DoEvents()

            Try
                Dim psi As New Diagnostics.ProcessStartInfo("python", $"""{scriptPath}"" --query ""{query}"" --json --max-pages 2") With {
                    .RedirectStandardOutput = True,
                    .RedirectStandardError  = True,
                    .UseShellExecute        = False,
                    .CreateNoWindow         = True
                }
                Using proc As Diagnostics.Process = Diagnostics.Process.Start(psi)
                    Dim output As String = proc.StandardOutput.ReadToEnd()
                    proc.WaitForExit(30000)

                    ' Parser les lignes JSON
                    Dim seen As New HashSet(Of String)()
                    For Each line As String In output.Split(vbLf)
                        Dim trimmed = line.Trim()
                        If String.IsNullOrEmpty(trimmed) Then Continue For
                        Try
                            Dim jo = Newtonsoft.Json.Linq.JObject.Parse(trimmed)
                            If jo("type")?.ToString() = "oeuvres" Then
                                Dim oeuvres = jo("data")
                                If oeuvres Is Nothing Then Continue For
                                For Each oe In oeuvres
                                    ' Extraire IPI de chaque ayant-droit
                                    For Each role In {"auteurs", "compositeurs", "editeurs", "arrangeurs", "adaptateurs"}
                                        Dim arr = oe(role)
                                        If arr Is Nothing Then Continue For
                                        For Each p In arr
                                            Dim ipi  As String = p("ipi")?.ToString().Trim()
                                            Dim pNom As String = p("nom")?.ToString().Trim()
                                            Dim pRole As String = p("role")?.ToString().Trim()
                                            If String.IsNullOrEmpty(ipi) OrElse String.IsNullOrEmpty(pNom) Then Continue For
                                            ' Filtrer par nom/prénom
                                            Dim pNomUp = pNom.ToUpper()
                                            Dim match = (String.IsNullOrEmpty(nom) OrElse pNomUp.Contains(nom.ToUpper())) AndAlso
                                                        (String.IsNullOrEmpty(prenom) OrElse pNomUp.Contains(prenom.ToUpper()))
                                            If match AndAlso Not seen.Contains(ipi) Then
                                                seen.Add(ipi)
                                                results.Add(Tuple.Create(ipi, pNom, If(pRole, "")))
                                            End If
                                        Next
                                    Next
                                Next
                            End If
                        Catch
                        End Try
                    Next
                End Using
            Catch ex As Exception
                errMsg = ex.Message
            End Try
            dlgWait.Close()
        End Using

        If errMsg <> "" Then
            MessageBox.Show("Erreur : " & errMsg, "Recherche SACEM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If results.Count = 0 Then
            MessageBox.Show("Aucun IPI trouvé pour : " & query, "Recherche SACEM", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Afficher le picker
        Using picker As New FormIPISACEMPicker(results)
            If picker.ShowDialog(Me) = DialogResult.OK Then
                For Each entry In picker.SelectedEntries
                    Dim nr As DataRow = _dtIPI.NewRow()
                    nr("Rôle(s)") = entry.Item3
                    nr("IPI")     = entry.Item1
                    nr("Nom")     = entry.Item2
                    _dtIPI.Rows.Add(nr)
                Next
            End If
        End Using
    End Sub

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
        ' IPI_LIST → peupler dgvIPI
        Dim ipiList As String = SafeStr(row, "IPI_LIST")
        If Not String.IsNullOrEmpty(ipiList) Then
            For Each entry In ipiList.Split(";"c)
                Dim parts = entry.Split("|"c)
                Dim nr As DataRow = _dtIPI.NewRow()
                nr("Rôle(s)") = If(parts.Length > 0, parts(0).Trim(), "")
                nr("IPI")     = If(parts.Length > 1, parts(1).Trim(), "")
                nr("Nom")     = If(parts.Length > 2, parts(2).Trim(), "")
                _dtIPI.Rows.Add(nr)
            Next
        End If
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

        ' ── IPI : valider chaque ligne du DGV ──
        Dim ipiEntries As New List(Of String)()
        For Each r2 As DataRow In _dtIPI.Rows
            Dim ipiVal As String = r2("IPI").ToString().Trim().Replace(" ", "")
            If String.IsNullOrEmpty(ipiVal) Then Continue For
            ipiVal = ipiVal.PadLeft(11, "0"c)
            If ipiVal.Length <> 11 OrElse Not ipiVal.All(Function(ch) Char.IsDigit(ch)) Then
                erreurs.Add($"• IPI ""{r2("IPI")}"" invalide (11 chiffres max).")
                Continue For
            End If
            Dim roles2 As String = r2("Rôle(s)").ToString().Trim().ToUpper()
            Dim nom2   As String = r2("Nom").ToString().Trim()
            ipiEntries.Add($"{roles2}|{ipiVal}|{nom2}")
        Next

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
        SetProprietes(nom, prenom, socCode, ipiEntries, coadVal, telVal, secuVal)

        Me.DialogResult = DialogResult.OK
    End Sub

    ' ── Setters (contournement ReadOnly via Friend) ───────────
    Private _pseudo    As String : Private _nom       As String
    Private _prenom    As String : Private _genre     As String
    Private _societe   As String : Private _ipiList   As String
    Private _coad      As String : Private _numVoie   As String
    Private _typeVoie  As String : Private _nomVoie   As String
    Private _cp        As String : Private _ville     As String
    Private _pays      As String : Private _mail      As String
    Private _tel       As String : Private _dateNaiss As String
    Private _lieuNaiss As String : Private _numSecu   As String


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
                               ipiEntries As List(Of String), coad As String,
                               tel As String, secu As String)
        _nom       = nom
        _prenom    = prenom
        _pseudo    = txtPseudo.Text.Trim()
        _genre     = If(cboGenre.SelectedItem?.ToString(), "").Trim()
        _societe   = societe
        _ipiList   = String.Join(";", ipiEntries)
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
    ''' <summary>Format : Roles|IPI|Nom;Roles|IPI|Nom</summary>
    Public ReadOnly Property ResultIPIList() As String
        Get
            Return _ipiList
        End Get
    End Property
    ''' <summary>Premier IPI seul — compat GetCOADIPI dans MainForm</summary>
    Public ReadOnly Property ResultIPI() As String
        Get
            If String.IsNullOrEmpty(_ipiList) Then Return ""
            Dim parts = _ipiList.Split(";"c)(0).Split("|"c)
            Return If(parts.Length > 1, parts(1).Trim(), "")
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

' ════════════════════════════════════════════════════════════════
' FormIPISACEMPicker — sélection IPI depuis résultats SACEM
' ════════════════════════════════════════════════════════════════
Public Class FormIPISACEMPicker
    Inherits Form

    Private _results  As List(Of Tuple(Of String, String, String))
    Private _clv      As ListView
    Public  SelectedEntries As New List(Of Tuple(Of String, String, String))()

    Public Sub New(results As List(Of Tuple(Of String, String, String)))
        _results = results
        InitUI()
    End Sub

    Private Sub InitUI()
        Me.Text            = "Sélectionner IPI(s) — Répertoire SACEM"
        Me.Size            = New Size(540, 420)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.BackColor       = Color.White

        Dim lbl As New Label() With {
            .Text     = "Cocher les IPI à importer :",
            .Location = New Point(10, 10),
            .Size     = New Size(500, 18),
            .Font     = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        Me.Controls.Add(lbl)

        _clv = New ListView() With {
            .Location      = New Point(10, 32),
            .Size          = New Size(504, 310),
            .View          = View.Details,
            .CheckBoxes    = True,
            .FullRowSelect  = True,
            .GridLines     = True,
            .Font          = New Font("Segoe UI", 9)
        }
        _clv.Columns.Add("IPI",    110)
        _clv.Columns.Add("Nom",    260)
        _clv.Columns.Add("Rôle",    90)

        For Each r In _results
            Dim li As New ListViewItem(r.Item1)
            li.SubItems.Add(r.Item2)
            li.SubItems.Add(r.Item3)
            _clv.Items.Add(li)
        Next
        Me.Controls.Add(_clv)

        Dim btnOK As New Button() With {
            .Text      = "Importer",
            .Location  = New Point(304, 350),
            .Size      = New Size(100, 28),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font      = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        Dim btnCancel As New Button() With {
            .Text      = "Annuler",
            .Location  = New Point(412, 350),
            .Size      = New Size(100, 28),
            .FlatStyle = FlatStyle.Flat
        }
        Me.Controls.Add(btnOK)
        Me.Controls.Add(btnCancel)

        AddHandler btnOK.Click, Sub(s, e)
            SelectedEntries.Clear()
            For Each li As ListViewItem In _clv.CheckedItems
                SelectedEntries.Add(Tuple.Create(li.Text, li.SubItems(1).Text, li.SubItems(2).Text))
            Next
            If SelectedEntries.Count = 0 Then
                MessageBox.Show("Cocher au moins un IPI.", "Sélection", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            Me.DialogResult = DialogResult.OK
        End Sub
        AddHandler btnCancel.Click, Sub(s, e) Me.DialogResult = DialogResult.Cancel

        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel
    End Sub
End Class
