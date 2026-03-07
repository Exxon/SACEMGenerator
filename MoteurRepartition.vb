Imports System.Data

''' <summary>
''' Moteur de calcul de répartition SACEM — Paroles + Musique.
''' Calcule automatiquement les parts DEP et DR de chaque ayant droit
''' selon les règles des Statuts et Règlement Général SACEM 2022.
'''
''' Règles implémentées :
'''   - Art. 57     : DEP tiers égaux auteur/compositeur/éditeur
'''   - Art. 58     : DEP inédit sans éditeur
'''   - Art. 66     : DEP avec adaptateur
'''   - Art. 70     : DEP avec arrangeur (et exception film/symphonique)
'''   - Art. 76     : DR protégé (avec/sans arrangeur/adaptateur)
'''   - Art. 77     : DR domaine public
'''   - Règle MIN   : Part DR éditeurs = MIN(50%, somme PH éditeurs)
'''                   Delta redistribué sur auteurs + compositeurs
'''   - Partage interne égalitaire ou inégalitaire (pondération PH)
''' </summary>
Public Class MoteurRepartition

    ' ─────────────────────────────────────────────────────────────
    ' TYPES
    ' ─────────────────────────────────────────────────────────────

    Public Enum TypeOeuvre
        ParolesEtMusique
        MusiqueSeule
        LitteraireSeule
    End Enum

    ''' <summary>Résultat de calcul pour un ayant droit.</summary>
    Public Class ResultatAyantDroit
        Public Designation As String
        Public Role As String           ' A, C, E, AR, AD
        Public ClePhono As Double       ' % saisi ou calculé
        Public PartDEP As Double        ' % calculé
        Public PartDR As Double         ' % calculé
        Public LettrageLie As String    ' lettre du créateur associé (pour éditeur)
    End Class

    ''' <summary>Paramètres d'entrée du moteur.</summary>
    Public Class ParamsOeuvre
        Public TypeOeuvre As TypeOeuvre = TypeOeuvre.ParolesEtMusique
        Public EstEditee As Boolean = True
        Public EstDomainePublic As Boolean = False
        Public EstFilmOuSymphonique As Boolean = False   ' exception arrangeur art.70/76
        Public Inegalitaire As Boolean = False
    End Class

    ' ─────────────────────────────────────────────────────────────
    ' POINT D'ENTRÉE PRINCIPAL
    ' ─────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Calcule les parts DEP et DR pour tous les ayants droit.
    ''' Met à jour directement les colonnes DE et DR du DataTable.
    ''' </summary>
    ''' <param name="dt">DataTable DtDepotCreateur</param>
    ''' <param name="params">Paramètres de l'œuvre</param>
    Public Shared Sub Calculer(dt As DataTable, params As ParamsOeuvre)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        ' 1. Recenser les ayants droit par catégorie
        Dim lignesA  As New List(Of DataRow)()   ' Auteurs
        Dim lignesC  As New List(Of DataRow)()   ' Compositeurs
        Dim lignesE  As New List(Of DataRow)()   ' Éditeurs
        Dim lignesAR As New List(Of DataRow)()   ' Arrangeurs
        Dim lignesAD As New List(Of DataRow)()   ' Adaptateurs

        For Each row As DataRow In dt.Rows
            Select Case row("Role").ToString().Trim().ToUpper()
                Case "A"  : lignesA.Add(row)
                Case "C"  : lignesC.Add(row)
                Case "E"  : lignesE.Add(row)
                Case "AR" : lignesAR.Add(row)
                Case "AD" : lignesAD.Add(row)
            End Select
        Next

        Dim nbA  As Integer = lignesA.Count
        Dim nbC  As Integer = lignesC.Count
        Dim nbE  As Integer = lignesE.Count
        Dim nbAR As Integer = lignesAR.Count
        Dim nbAD As Integer = lignesAD.Count

        Dim aArrangeur  As Boolean = nbAR > 0
        Dim aAdaptateur As Boolean = nbAD > 0

        ' 2. Somme des PH éditeurs (pour règle DR MIN 50%)
        Dim sommePHEditeurs As Double = 0
        For Each row As DataRow In lignesE
            sommePHEditeurs += ParsePH(row)
        Next

        ' 3. Calculer parts catégories DEP
        Dim depA  As Double = 0
        Dim depC  As Double = 0
        Dim depE  As Double = 0
        Dim depAR As Double = 0
        Dim depAD As Double = 0

        CalculerPartsDEP(params, nbA, nbC, nbE, aArrangeur, aAdaptateur,
                         depA, depC, depE, depAR, depAD)

        ' 4. Calculer parts catégories DR
        Dim drA  As Double = 0
        Dim drC  As Double = 0
        Dim drE  As Double = 0
        Dim drAR As Double = 0
        Dim drAD As Double = 0

        CalculerPartsDR(params, nbA, nbC, nbE, aArrangeur, aAdaptateur,
                        sommePHEditeurs,
                        drA, drC, drE, drAR, drAD)

        ' 5. Répartir en interne par catégorie et écrire dans le DataTable
        RepartirCategorie(lignesA,  depA,  drA,  params.Inegalitaire, dt)
        RepartirCategorie(lignesC,  depC,  drC,  params.Inegalitaire, dt)
        RepartirEditeurs (lignesE,  depE,  drE,  sommePHEditeurs, params.Inegalitaire, dt)
        RepartirCategorie(lignesAR, depAR, drAR, params.Inegalitaire, dt)
        RepartirCategorie(lignesAD, depAD, drAD, params.Inegalitaire, dt)
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CALCUL PARTS DEP PAR CATÉGORIE
    ' ─────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Calcule les parts globales DEP par catégorie.
    ''' Sources : Art. 57, 58, 66, 70 RG SACEM 2022.
    ''' </summary>
    Private Shared Sub CalculerPartsDEP(params As ParamsOeuvre,
                                         nbA As Integer, nbC As Integer, nbE As Integer,
                                         aArrangeur As Boolean, aAdaptateur As Boolean,
                                         ByRef depA As Double, ByRef depC As Double,
                                         ByRef depE As Double, ByRef depAR As Double,
                                         ByRef depAD As Double)
        Dim aEditeur As Boolean = nbE > 0

        If params.EstDomainePublic Then
            ' Domaine public — arrangeur et/ou adaptateur uniquement
            CalculerPartsDEPDomainePublic(aEditeur, aArrangeur, aAdaptateur,
                                          depA, depC, depE, depAR, depAD)
            Return
        End If

        Dim aA As Boolean = nbA > 0
        Dim aC As Boolean = nbC > 0
        Dim nbCreateursCats As Integer = If(aA, 1, 0) + If(aC, 1, 0)

        ' Part de base par catégorie créateur
        ' Éditée  : total créateurs = 2/3, éditeur = 1/3
        ' Inédite : total créateurs = 100%, éditeur = 0
        Dim totalCreateurs As Double = If(aEditeur, 200.0 / 3.0, 100.0)
        Dim partParCreateur As Double = If(nbCreateursCats > 0, totalCreateurs / nbCreateursCats, 0)
        Dim partEditeur As Double = If(aEditeur, 100.0 / 3.0, 0)

        If Not aArrangeur AndAlso Not aAdaptateur Then
            depA  = If(aA, partParCreateur, 0)
            depC  = If(aC, partParCreateur, 0)
            depE  = partEditeur

        ElseIf aArrangeur AndAlso Not aAdaptateur Then
            ' Arrangeur prend 2/24 sur le total
            ' Exception film/symph : arrangeur prend 4/24
            Dim fracAR As Double = If(params.EstFilmOuSymphonique, 4.0 / 24.0, 2.0 / 24.0)
            Dim reductionParCreateur As Double = If(nbCreateursCats > 0, fracAR * 100.0 / nbCreateursCats, 0)
            depA  = If(aA, partParCreateur - reductionParCreateur, 0)
            depC  = If(aC, partParCreateur - reductionParCreateur, 0)
            depE  = partEditeur
            depAR = fracAR * 100.0

        ElseIf Not aArrangeur AndAlso aAdaptateur Then
            ' Adaptateur prend 2/24 sur les auteurs
            Dim fracAD As Double = 2.0 / 24.0
            Dim reductionParCreateur As Double = If(nbCreateursCats > 0, fracAD * 100.0 / nbCreateursCats, 0)
            depA  = If(aA, partParCreateur - reductionParCreateur, 0)
            depC  = If(aC, partParCreateur - reductionParCreateur, 0)
            depE  = partEditeur
            depAD = fracAD * 100.0

        ElseIf aArrangeur AndAlso aAdaptateur Then
            ' Arrangeur 2/24 + adaptateur 4/24
            Dim fracAR As Double = If(params.EstFilmOuSymphonique, 4.0 / 24.0, 2.0 / 24.0)
            Dim fracAD As Double = 4.0 / 24.0
            Dim reductionParCreateur As Double = If(nbCreateursCats > 0, (fracAR + fracAD) * 100.0 / nbCreateursCats, 0)
            depA  = If(aA, partParCreateur - reductionParCreateur, 0)
            depC  = If(aC, partParCreateur - reductionParCreateur, 0)
            depE  = partEditeur
            depAR = fracAR * 100.0
            depAD = fracAD * 100.0
        End If
    End Sub

    ''' <summary>Parts DEP domaine public — Art. 77.</summary>
    Private Shared Sub CalculerPartsDEPDomainePublic(aEditeur As Boolean,
                                                      aArrangeur As Boolean,
                                                      aAdaptateur As Boolean,
                                                      ByRef depA As Double,
                                                      ByRef depC As Double,
                                                      ByRef depE As Double,
                                                      ByRef depAR As Double,
                                                      ByRef depAD As Double)
        If aArrangeur AndAlso aAdaptateur Then
            If aEditeur Then
                depE  = 50.0 : depAR = 25.0 : depAD = 25.0
            Else
                depAR = 50.0 : depAD = 50.0
            End If
        ElseIf aArrangeur Then
            If aEditeur Then
                depE = 50.0 : depAR = 50.0
            Else
                depAR = 100.0
            End If
        ElseIf aAdaptateur Then
            If aEditeur Then
                depE = 50.0 : depAD = 50.0
            Else
                depAD = 100.0
            End If
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CALCUL PARTS DR PAR CATÉGORIE
    ' ─────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Calcule les parts globales DR par catégorie.
    ''' Sources : Art. 76, 77 RG SACEM 2022.
    ''' Règle MIN : Part DR éditeurs = MIN(50%, sommePHEditeurs)
    '''             Delta redistribué également sur A et C.
    ''' </summary>
    Private Shared Sub CalculerPartsDR(params As ParamsOeuvre,
                                        nbA As Integer, nbC As Integer, nbE As Integer,
                                        aArrangeur As Boolean, aAdaptateur As Boolean,
                                        sommePHEditeurs As Double,
                                        ByRef drA As Double, ByRef drC As Double,
                                        ByRef drE As Double, ByRef drAR As Double,
                                        ByRef drAD As Double)
        Dim aEditeur As Boolean = nbE > 0

        If params.EstDomainePublic Then
            CalculerPartsDRDomainePublic(aEditeur, aArrangeur, aAdaptateur,
                                         sommePHEditeurs,
                                         drA, drC, drE, drAR, drAD)
            Return
        End If

        ' Parts théoriques à 50% éditeur (Art. 76)
        Dim drA_theorique  As Double = 0
        Dim drC_theorique  As Double = 0
        Dim drE_theorique  As Double = 0
        Dim drAR_theorique As Double = 0
        Dim drAD_theorique As Double = 0

        ' Parts créateurs selon présence A et/ou C
        Dim aA As Boolean = nbA > 0
        Dim aC As Boolean = nbC > 0

        ' Part de base pour A et C selon qui est présent (hors arrangeur/adaptateur)
        ' Avec éditeur : total créateurs = 50%, sans éditeur : total créateurs = 100%
        Dim totalCreateurs As Double = If(aEditeur, 50.0, 100.0)
        Dim nbCreateursCats As Integer = If(aA, 1, 0) + If(aC, 1, 0)
        Dim partParCreateur As Double = If(nbCreateursCats > 0, totalCreateurs / nbCreateursCats, 0)

        If Not aArrangeur AndAlso Not aAdaptateur Then
            ' CAS 1 & 2
            drA_theorique  = If(aA, partParCreateur, 0)
            drC_theorique  = If(aC, partParCreateur, 0)
            drE_theorique  = If(aEditeur, 50.0, 0)

        ElseIf aArrangeur AndAlso Not aAdaptateur Then
            ' CAS 3 & 4 — arrangeur prend 6.25% (ou 12.5% film/symph) sur chaque créateur présent
            If params.EstFilmOuSymphonique Then
                Dim partAR As Double = 12.5
                Dim reductionParCreateur As Double = If(nbCreateursCats > 0, partAR / nbCreateursCats, 0)
                drA_theorique  = If(aA, partParCreateur - reductionParCreateur, 0)
                drC_theorique  = If(aC, partParCreateur - reductionParCreateur, 0)
                drE_theorique  = If(aEditeur, 50.0, 0)
                drAR_theorique = partAR
            Else
                Dim partAR As Double = 6.25
                Dim reductionParCreateur As Double = If(nbCreateursCats > 0, partAR / nbCreateursCats, 0)
                drA_theorique  = If(aA, partParCreateur - reductionParCreateur, 0)
                drC_theorique  = If(aC, partParCreateur - reductionParCreateur, 0)
                drE_theorique  = If(aEditeur, 50.0, 0)
                drAR_theorique = partAR
            End If

        ElseIf Not aArrangeur AndAlso aAdaptateur Then
            ' CAS 5 & 6 — adaptateur prend 12.5% sur les auteurs uniquement
            Dim partAD As Double = 12.5
            Dim reductionParCreateur As Double = If(nbCreateursCats > 0, partAD / nbCreateursCats, 0)
            drA_theorique  = If(aA, partParCreateur - reductionParCreateur, 0)
            drC_theorique  = If(aC, partParCreateur - reductionParCreateur, 0)
            drE_theorique  = If(aEditeur, 50.0, 0)
            drAD_theorique = partAD

        ElseIf aArrangeur AndAlso aAdaptateur Then
            ' CAS 7 — arrangeur 6.25% + adaptateur 12.5%
            Dim partAR As Double = 6.25
            Dim partAD As Double = 12.5
            Dim reductionParCreateur As Double = If(nbCreateursCats > 0, (partAR + partAD) / nbCreateursCats, 0)
            drA_theorique  = If(aA, partParCreateur - reductionParCreateur, 0)
            drC_theorique  = If(aC, partParCreateur - reductionParCreateur, 0)
            drE_theorique  = If(aEditeur, 50.0, 0)
            drAR_theorique = partAR
            drAD_theorique = partAD
        End If

        ' Appliquer règle MIN(50%, sommePHEditeurs) sur la part éditeur
        drAR = drAR_theorique
        drAD = drAD_theorique

        If aEditeur Then
            Dim partEEffective As Double = Math.Min(drE_theorique, sommePHEditeurs)
            Dim delta As Double = drE_theorique - partEEffective

            drE = partEEffective

            ' Redistribuer le delta sur A et C
            Dim nbCreateurs As Integer = If(nbA > 0, 1, 0) + If(nbC > 0, 1, 0)
            If nbCreateurs > 0 Then
                Dim deltaParCreateur As Double = delta / nbCreateurs
                drA = drA_theorique + If(nbA > 0, deltaParCreateur, 0)
                drC = drC_theorique + If(nbC > 0, deltaParCreateur, 0)
            Else
                drA = drA_theorique
                drC = drC_theorique
            End If
        Else
            drA = drA_theorique
            drC = drC_theorique
            drE = 0
        End If
    End Sub

    ''' <summary>Parts DR domaine public — Art. 77 avec règle MIN éditeur.</summary>
    Private Shared Sub CalculerPartsDRDomainePublic(aEditeur As Boolean,
                                                     aArrangeur As Boolean,
                                                     aAdaptateur As Boolean,
                                                     sommePHEditeurs As Double,
                                                     ByRef drA As Double,
                                                     ByRef drC As Double,
                                                     ByRef drE As Double,
                                                     ByRef drAR As Double,
                                                     ByRef drAD As Double)
        Dim drE_theorique As Double = 0
        Dim drAR_theorique As Double = 0
        Dim drAD_theorique As Double = 0

        If aArrangeur AndAlso aAdaptateur Then
            If aEditeur Then
                drE_theorique = 50.0 : drAR_theorique = 25.0 : drAD_theorique = 25.0
            Else
                drAR_theorique = 50.0 : drAD_theorique = 50.0
            End If
        ElseIf aArrangeur Then
            If aEditeur Then
                drE_theorique = 50.0 : drAR_theorique = 50.0
            Else
                drAR_theorique = 100.0
            End If
        ElseIf aAdaptateur Then
            If aEditeur Then
                drE_theorique = 50.0 : drAD_theorique = 50.0
            Else
                drAD_theorique = 100.0
            End If
        End If

        drAR = drAR_theorique
        drAD = drAD_theorique

        If aEditeur Then
            Dim partEEffective As Double = Math.Min(drE_theorique, sommePHEditeurs)
            drE = partEEffective
            ' Pas de redistribution delta en domaine public (pas de A/C originaux)
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' RÉPARTITION INTERNE PAR CATÉGORIE
    ' ─────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Répartit la part globale d'une catégorie entre ses membres.
    ''' Égalitaire : division égale.
    ''' Inégalitaire : pondération par clé PHONO individuelle.
    ''' </summary>
    Private Shared Sub RepartirCategorie(lignes As List(Of DataRow),
                                          partGlobaleDEP As Double,
                                          partGlobaleDR As Double,
                                          inegalitaire As Boolean,
                                          dt As DataTable)
        If lignes.Count = 0 Then Return

        If Not inegalitaire OrElse lignes.Count = 1 Then
            ' Égalitaire
            Dim depIndiv As Double = Math.Round(partGlobaleDEP / lignes.Count, 4)
            Dim drIndiv  As Double = Math.Round(partGlobaleDR  / lignes.Count, 4)
            For Each row As DataRow In lignes
                EcrirePartsRow(row, depIndiv, drIndiv, dt)
            Next
        Else
            ' Inégalitaire — pondération par PH
            Dim sommePH As Double = lignes.Sum(Function(r) ParsePH(r))
            If sommePH = 0 Then
                ' Fallback égalitaire si PH tous à 0
                RepartirCategorie(lignes, partGlobaleDEP, partGlobaleDR, False, dt)
                Return
            End If
            For Each row As DataRow In lignes
                Dim ph As Double = ParsePH(row)
                Dim depIndiv As Double = Math.Round(partGlobaleDEP * ph / sommePH, 4)
                Dim drIndiv  As Double = Math.Round(partGlobaleDR  * ph / sommePH, 4)
                EcrirePartsRow(row, depIndiv, drIndiv, dt)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Répartition spéciale pour les éditeurs.
    ''' La part DR effective est déjà calculée (MIN 50% appliqué).
    ''' En inégalitaire : pondération par PH individuel / sommePHEditeurs.
    ''' </summary>
    Private Shared Sub RepartirEditeurs(lignesE As List(Of DataRow),
                                         partGlobaleDEP As Double,
                                         partGlobaleDR As Double,
                                         sommePHEditeurs As Double,
                                         inegalitaire As Boolean,
                                         dt As DataTable)
        If lignesE.Count = 0 Then Return

        If Not inegalitaire OrElse lignesE.Count = 1 Then
            ' Égalitaire entre éditeurs
            Dim depIndiv As Double = Math.Round(partGlobaleDEP / lignesE.Count, 4)
            Dim drIndiv  As Double = Math.Round(partGlobaleDR  / lignesE.Count, 4)
            For Each row As DataRow In lignesE
                EcrirePartsRow(row, depIndiv, drIndiv, dt)
            Next
        Else
            ' Inégalitaire — pondération par PH individuel / sommePHEditeurs
            If sommePHEditeurs = 0 Then
                RepartirEditeurs(lignesE, partGlobaleDEP, partGlobaleDR, 0, False, dt)
                Return
            End If
            For Each row As DataRow In lignesE
                Dim ph As Double = ParsePH(row)
                Dim depIndiv As Double = Math.Round(partGlobaleDEP * ph / sommePHEditeurs, 4)
                Dim drIndiv  As Double = Math.Round(partGlobaleDR  * ph / sommePHEditeurs, 4)
                EcrirePartsRow(row, depIndiv, drIndiv, dt)
            Next
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' CALCUL PH PAR DÉFAUT LORS DE L'AJOUT
    ' ─────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Calcule la clé PHONO par défaut d'un nouvel ayant droit lors de son ajout.
    ''' Par défaut : PH = part DR qui sera la sienne après recalcul.
    ''' Conservation des proportions pour les éditeurs existants.
    ''' </summary>
    ''' <param name="dt">DataTable AVANT ajout du nouvel ayant droit</param>
    ''' <param name="nouveauRole">Rôle du nouvel ayant droit (A, C, E, AR, AD)</param>
    ''' <param name="lettrageParent">Pour un éditeur : lettre du créateur associé</param>
    ''' <param name="params">Paramètres de l'œuvre</param>
    ''' <returns>Clé PHONO initiale suggérée (0-100)</returns>
    Public Shared Function CalculerPHDefaut(dt As DataTable,
                                             nouveauRole As String,
                                             lettrageParent As String,
                                             params As ParamsOeuvre) As Double
        ' Simuler l'ajout et recalculer
        Dim nbA  As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "A")
        Dim nbC  As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "C")
        Dim nbE  As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "E")
        Dim nbAR As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "AR")
        Dim nbAD As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "AD")

        ' Incrémenter selon le rôle ajouté
        Select Case nouveauRole.ToUpper()
            Case "A"  : nbA  += 1
            Case "C"  : nbC  += 1
            Case "E"  : nbE  += 1
            Case "AR" : nbAR += 1
            Case "AD" : nbAD += 1
        End Select

        Dim aArrangeur  As Boolean = nbAR > 0
        Dim aAdaptateur As Boolean = nbAD > 0

        ' Calculer parts DR après ajout
        Dim sommePHEditeurs As Double = dt.AsEnumerable().
            Where(Function(r) r("Role").ToString() = "E").
            Sum(Function(r) ParsePH(r))

        Dim drA As Double = 0, drC As Double = 0, drE As Double = 0
        Dim drAR As Double = 0, drAD As Double = 0

        CalculerPartsDR(params, nbA, nbC, nbE, aArrangeur, aAdaptateur,
                        sommePHEditeurs, drA, drC, drE, drAR, drAD)

        ' Retourner la part DR de la catégorie du nouvel ayant droit
        Select Case nouveauRole.ToUpper()
            Case "A"  : Return If(nbA > 0, Math.Round(drA / nbA, 4), 0)
            Case "C"  : Return If(nbC > 0, Math.Round(drC / nbC, 4), 0)
            Case "AR" : Return If(nbAR > 0, Math.Round(drAR / nbAR, 4), 0)
            Case "AD" : Return If(nbAD > 0, Math.Round(drAD / nbAD, 4), 0)
            Case "E"
                ' Pour un éditeur : PH = part DR de son créateur associé / nb éditeurs de ce créateur
                If Not String.IsNullOrEmpty(lettrageParent) Then
                    Dim editeursDuCreateur As Integer = dt.AsEnumerable().
                        Count(Function(r) r("Role").ToString() = "E" AndAlso
                                          r("Lettrage").ToString() = lettrageParent) + 1
                    Dim partDRCreateur As Double = 0
                    Dim creaRow As DataRow = dt.AsEnumerable().
                        FirstOrDefault(Function(r) (r("Role").ToString() = "A" OrElse
                                                    r("Role").ToString() = "C") AndAlso
                                                    r("Lettrage").ToString() = lettrageParent)
                    If creaRow IsNot Nothing Then
                        Select Case creaRow("Role").ToString()
                            Case "A" : partDRCreateur = If(nbA > 0, Math.Round(drA / nbA, 4), 0)
                            Case "C" : partDRCreateur = If(nbC > 0, Math.Round(drC / nbC, 4), 0)
                        End Select
                    End If
                    Return Math.Round(partDRCreateur / editeursDuCreateur, 4)
                End If
                Return If(nbE > 0, Math.Round(drE / nbE, 4), 0)
        End Select

        Return 0
    End Function

    ''' <summary>
    ''' Recalcule et redistribue les PH de tous les ayants droit après ajout/suppression.
    ''' Conservation des proportions internes aux coéditeurs d'un même créateur.
    ''' Principe : PH créateur = part DR théorique (sans éditeur), PH éditeur = fraction de ce PH.
    ''' </summary>
    Public Shared Sub RecalculerPHApresAjout(dt As DataTable, params As ParamsOeuvre)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim nbA  As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "A")
        Dim nbC  As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "C")
        Dim nbE  As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "E")
        Dim nbAR As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "AR")
        Dim nbAD As Integer = dt.AsEnumerable().Count(Function(r) r("Role").ToString() = "AD")

        ' ── Étape 1 : calculer les parts DR théoriques à 50% éditeur ──────────
        ' On passe sommePHEditeurs = 50 pour forcer le calcul au plafond théorique
        ' Cela évite le cercle vicieux PH éditeur → DR → PH éditeur
        Dim drA As Double = 0, drC As Double = 0, drE As Double = 0
        Dim drAR As Double = 0, drAD As Double = 0

        CalculerPartsDR(params, nbA, nbC, nbE, nbAR > 0, nbAD > 0,
                        50.0,   ' sommePHEditeurs forcée à 50 → plafond théorique
                        drA, drC, drE, drAR, drAD)

        ' ── Étape 2 : PH créateurs = leur part DR théorique individuelle ───────
        Dim lignesA  = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "A").ToList()
        Dim lignesC  = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "C").ToList()
        Dim lignesAR = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "AR").ToList()
        Dim lignesAD = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "AD").ToList()

        For Each row As DataRow In lignesA
            row("PH") = Math.Round(drA / nbA, 4).ToString(Globalization.CultureInfo.InvariantCulture)
        Next
        For Each row As DataRow In lignesC
            row("PH") = Math.Round(drC / nbC, 4).ToString(Globalization.CultureInfo.InvariantCulture)
        Next
        For Each row As DataRow In lignesAR
            If nbAR > 0 Then row("PH") = Math.Round(drAR / nbAR, 4).ToString(Globalization.CultureInfo.InvariantCulture)
        Next
        For Each row As DataRow In lignesAD
            If nbAD > 0 Then row("PH") = Math.Round(drAD / nbAD, 4).ToString(Globalization.CultureInfo.InvariantCulture)
        Next

        ' ── Étape 3 : PH éditeurs = fraction du PH de leur créateur associé ───
        ' Conservation des proportions relatives entre coéditeurs d'un même créateur
        Dim lignesE = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "E").ToList()
        Dim groupesEditeurs = lignesE.GroupBy(Function(r) r("Lettrage").ToString())

        For Each groupe In groupesEditeurs
            Dim lettrage As String = groupe.Key
            Dim editeursDuGroupe = groupe.ToList()

            ' Trouver le créateur associé et sa part DR théorique
            Dim creaRow As DataRow = dt.AsEnumerable().
                FirstOrDefault(Function(r) (r("Role").ToString() = "A" OrElse
                                            r("Role").ToString() = "C") AndAlso
                                            r("Lettrage").ToString() = lettrage)

            Dim partDRCrea As Double = 0
            If creaRow IsNot Nothing Then
                Select Case creaRow("Role").ToString()
                    Case "A" : partDRCrea = If(nbA > 0, drA / nbA, 0)
                    Case "C" : partDRCrea = If(nbC > 0, drC / nbC, 0)
                End Select
            End If

            ' Lire les PH provisoires des éditeurs du groupe (parts relatives entre eux)
            Dim sommePHGroupe As Double = editeursDuGroupe.Sum(Function(r) ParsePH(r))

            If sommePHGroupe = 0 Then
                ' Pas de PH provisoire → répartition égale
                Dim phIndiv As Double = Math.Round(partDRCrea / editeursDuGroupe.Count, 4)
                For Each row As DataRow In editeursDuGroupe
                    row("PH") = phIndiv.ToString(Globalization.CultureInfo.InvariantCulture)
                Next
            Else
                ' Conserver les proportions relatives et ramener à la part DR du créateur
                For Each row As DataRow In editeursDuGroupe
                    Dim ph As Double = ParsePH(row)
                    Dim nouvPH As Double = Math.Round(partDRCrea * ph / sommePHGroupe, 4)
                    row("PH") = nouvPH.ToString(Globalization.CultureInfo.InvariantCulture)
                Next
            End If
        Next
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' UTILITAIRES
    ' ─────────────────────────────────────────────────────────────

    Private Shared Sub EcrirePartsRow(row As DataRow, dep As Double, dr As Double, dt As DataTable)
        If dt.Columns.Contains("DE") Then
            row("DE") = dep.ToString("0.####", Globalization.CultureInfo.InvariantCulture)
        End If
        If dt.Columns.Contains("DR") Then
            row("DR") = dr.ToString("0.####", Globalization.CultureInfo.InvariantCulture)
        End If
    End Sub

    Public Shared Function ParsePH(row As DataRow) As Double
        Dim ph As Double = 0
        Double.TryParse(row("PH").ToString().Replace(",", "."),
                        Globalization.NumberStyles.Any,
                        Globalization.CultureInfo.InvariantCulture, ph)
        Return ph
    End Function

End Class
