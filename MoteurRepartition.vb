Imports System.Data

''' <summary>
''' Moteur de calcul de répartition SACEM — Paroles + Musique.
''' Calcule automatiquement les parts DEP et DR de chaque ayant droit
''' selon les règles des Statuts et Règlement Général SACEM 2022.
'''
''' Règles implémentées :
'''   - Art. 57     : DEP tiers égaux auteur/compositeur/éditeur
'''   - Art. 58     : DEP inédit (part inédite INDIVIDUELLE par créateur)
'''   - Art. 66     : DEP avec adaptateur
'''   - Art. 70     : DEP avec arrangeur (et exception film/symphonique)
'''   - Art. 76     : DR protégé (avec/sans arrangeur/adaptateur)
'''   - Art. 77     : DR domaine public
'''   - Règle MIN   : Part DR éditeur = MIN(part statutaire, somme PH éditeurs du groupe)
'''                   Delta redistribué AU créateur du même lettrage (pas globalement)
'''   - Part inédite : créateur sans éditeur = part créateur + part éditeur imputable
'''                    (calculé par lettrage, indépendamment des autres créateurs)
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
    '''
    ''' RÈGLE PART INÉDITE (par créateur) :
    '''   Si un A ou C n'a pas d'éditeur associé (aucun E avec son lettrage),
    '''   il récupère sa propre part d'éditeur statutaire en plus de sa part créateur.
    '''   Ce calcul est fait INDIVIDUELLEMENT par lettrage, pas globalement.
    '''
    '''   Exemple : 1A sans éditeur + 1C avec éditeur
    '''     Part statutaire : A=33.333, C=33.333, E=33.333
    '''     → A (inédit) récupère la part éditeur qui lui est imputable = 33.333/2 * 1
    '''       (le 1/3 éditeur se répartit entre les créateurs proportionnellement)
    '''     Résultat : A DEP=50, C DEP=33.333, E DEP=16.667
    '''
    '''   De même pour DR : le delta MIN(50%, PHEditeur) s'applique par groupe lettrage.
    ''' </summary>
    Public Shared Sub Calculer(dt As DataTable, params As ParamsOeuvre)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        ' 1. Recenser les ayants droit par catégorie
        Dim lignesA  As New List(Of DataRow)()
        Dim lignesC  As New List(Of DataRow)()
        Dim lignesE  As New List(Of DataRow)()
        Dim lignesAR As New List(Of DataRow)()
        Dim lignesAD As New List(Of DataRow)()

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

        ' 2. Identifier quels lettrages ont un éditeur
        Dim lettragesAvecEditeur As New HashSet(Of String)(
            lignesE.Select(Function(r) r("Lettrage").ToString().Trim()))

        ' 3. Calculer les parts statutaires globales (comme si tout était édité)
        '    nbE fictif = nbA + nbC pour le calcul statutaire de base
        Dim nbA_fict  As Integer = nbA
        Dim nbC_fict  As Integer = nbC
        Dim nbE_fict  As Integer = nbA + nbC  ' 1 éditeur fictif par créateur pour les parts statutaires

        ' Parts statutaires DEP : A=1/3, C=1/3, E=1/3 (si A+C présents)
        ' On calcule la part éditeur imputable à chaque catégorie créateur
        Dim nbCreateursCats As Integer = If(nbA > 0, 1, 0) + If(nbC > 0, 1, 0)
        Dim partStatutaireEditeurTotal As Double = If(nbCreateursCats > 0 AndAlso nbE > 0, 100.0 / 3.0, 0)
        ' Part éditeur imputable à chaque catégorie créateur (A ou C) = partEditeurTotal / nbCreateursCats
        Dim partEditeurParCatCreateur As Double = If(nbCreateursCats > 0, partStatutaireEditeurTotal / nbCreateursCats, 0)

        ' Parts statutaires DR : A=25%, C=25%, E=50% (avec A+C)
        Dim partStatutaireDRETotal As Double = If(nbE > 0, 50.0, 0)
        Dim partDREditeurParCatCreateur As Double = If(nbCreateursCats > 0, partStatutaireDRETotal / nbCreateursCats, 0)

        ' 4. Calculer parts DEP/DR effectives par créateur individuel
        '    Créateur sans éditeur → récupère sa part + part éditeur imputable
        '    Créateur avec éditeur → part normale, éditeur reçoit sa part
        Dim depParCreateurA As New Dictionary(Of DataRow, Double)()
        Dim depParCreateurC As New Dictionary(Of DataRow, Double)()
        Dim drParCreateurA  As New Dictionary(Of DataRow, Double)()
        Dim drParCreateurC  As New Dictionary(Of DataRow, Double)()

        ' Réductions sur la part créateur dues aux arrangeurs/adaptateurs
        Dim reductionDEP_A As Double = 0, reductionDEP_C As Double = 0
        Dim reductionDR_A  As Double = 0, reductionDR_C  As Double = 0

        ' Ajustements arrangeur/adaptateur DEP
        Dim depAR As Double = 0, depAD As Double = 0
        Dim drAR  As Double = 0, drAD  As Double = 0

        If Not params.EstDomainePublic Then
            If aArrangeur Then
                Dim fracAR_dep As Double = If(params.EstFilmOuSymphonique, 4.0 / 24.0, 2.0 / 24.0)
                depAR = fracAR_dep * 100.0
                Dim redAR As Double = If(nbCreateursCats > 0, depAR / nbCreateursCats, 0)
                reductionDEP_A += If(nbA > 0, redAR, 0)
                reductionDEP_C += If(nbC > 0, redAR, 0)
                Dim fracAR_dr As Double = If(params.EstFilmOuSymphonique, 12.5, 6.25)
                drAR = fracAR_dr
                Dim redDR_AR As Double = If(nbCreateursCats > 0, drAR / nbCreateursCats, 0)
                reductionDR_A += If(nbA > 0, redDR_AR, 0)
                reductionDR_C += If(nbC > 0, redDR_AR, 0)
            End If
            If aAdaptateur Then
                Dim fracAD_dep As Double = 2.0 / 24.0
                depAD = fracAD_dep * 100.0
                Dim redAD As Double = If(nbCreateursCats > 0, depAD / nbCreateursCats, 0)
                reductionDEP_A += If(nbA > 0, redAD, 0)
                reductionDEP_C += If(nbC > 0, redAD, 0)
                Dim fracAD_dr As Double = 12.5
                drAD = fracAD_dr
                Dim redDR_AD As Double = If(nbCreateursCats > 0, drAD / nbCreateursCats, 0)
                reductionDR_A += If(nbA > 0, redDR_AD, 0)
                reductionDR_C += If(nbC > 0, redDR_AD, 0)
            End If
        End If

        ' 5. Appliquer part inédite individuelle par créateur
        '    Formule : part_individuelle = part_totale_categorie / (nbCats × nbMembres)
        '    Bonus inédit = part_ed_totale / (nbCats × nbMembres_catégorie)
        '    (prorata de la part éditeur imputable à ce créateur spécifiquement)
        For Each row As DataRow In lignesA
            Dim lettrage As String = row("Lettrage").ToString().Trim()
            Dim aEditeurIndiv As Boolean = lettragesAvecEditeur.Contains(lettrage)
            Dim depIndiv As Double = (100.0 - partStatutaireEditeurTotal) / (nbCreateursCats * nbA) - reductionDEP_A / nbA
            Dim drIndiv  As Double = (100.0 - partStatutaireDRETotal)     / (nbCreateursCats * nbA) - reductionDR_A  / nbA
            Dim bonusDEP As Double = If(Not aEditeurIndiv, partStatutaireEditeurTotal / (nbCreateursCats * nbA), 0)
            Dim bonusDR  As Double = If(Not aEditeurIndiv, partStatutaireDRETotal     / (nbCreateursCats * nbA), 0)
            depParCreateurA(row) = depIndiv + bonusDEP
            drParCreateurA(row)  = drIndiv  + bonusDR
        Next

        For Each row As DataRow In lignesC
            Dim lettrage As String = row("Lettrage").ToString().Trim()
            Dim aEditeurIndiv As Boolean = lettragesAvecEditeur.Contains(lettrage)
            Dim depIndiv As Double = (100.0 - partStatutaireEditeurTotal) / (nbCreateursCats * nbC) - reductionDEP_C / nbC
            Dim drIndiv  As Double = (100.0 - partStatutaireDRETotal)     / (nbCreateursCats * nbC) - reductionDR_C  / nbC
            Dim bonusDEP As Double = If(Not aEditeurIndiv, partStatutaireEditeurTotal / (nbCreateursCats * nbC), 0)
            Dim bonusDR  As Double = If(Not aEditeurIndiv, partStatutaireDRETotal     / (nbCreateursCats * nbC), 0)
            depParCreateurC(row) = depIndiv + bonusDEP
            drParCreateurC(row)  = drIndiv  + bonusDR
        Next

        ' 6. Pour les éditeurs : DEP/DR = part imputable au créateur de leur groupe
        '    Part éditeur du groupe = part_ed_totale / (nbCats × nb_membres_catégorie_du_créateur)
        '    DR éditeur : règle MIN(part statutaire groupe, sommePH éditeurs du groupe)
        '                 delta DR → redistribué au créateur du même lettrage
        Dim depEditeurParGroupe As New Dictionary(Of String, Double)()
        Dim drEditeurParGroupe  As New Dictionary(Of String, Double)()

        Dim lettragesEditeurs = lignesE.Select(Function(r) r("Lettrage").ToString().Trim()).Distinct().ToList()
        For Each lettrage As String In lettragesEditeurs
            Dim edsDuGroupe = lignesE.Where(Function(r) r("Lettrage").ToString().Trim() = lettrage).ToList()
            Dim sommePHGroupe As Double = edsDuGroupe.Sum(Function(r) ParsePH(r))

            ' Identifier la catégorie et nb membres du créateur associé
            Dim nbMembresCatCrea As Integer = 1
            For Each r As DataRow In lignesA
                If r("Lettrage").ToString().Trim() = lettrage Then nbMembresCatCrea = nbA : Exit For
            Next
            For Each r As DataRow In lignesC
                If r("Lettrage").ToString().Trim() = lettrage Then nbMembresCatCrea = nbC : Exit For
            Next

            Dim depGroupe As Double = partStatutaireEditeurTotal / (nbCreateursCats * nbMembresCatCrea)
            Dim drStat    As Double = partStatutaireDRETotal      / (nbCreateursCats * nbMembresCatCrea)
            Dim drEffectif As Double = Math.Min(drStat, sommePHGroupe)
            Dim deltaGroupe As Double = drStat - drEffectif

            depEditeurParGroupe(lettrage) = depGroupe
            drEditeurParGroupe(lettrage)  = drEffectif

            ' Redistribuer delta DR au créateur du groupe (par lettrage = 1 créateur spécifique)
            If deltaGroupe > 0 Then
                For Each r As DataRow In lignesA
                    If r("Lettrage").ToString().Trim() = lettrage Then
                        drParCreateurA(r) = drParCreateurA(r) + deltaGroupe
                    End If
                Next
                For Each r As DataRow In lignesC
                    If r("Lettrage").ToString().Trim() = lettrage Then
                        drParCreateurC(r) = drParCreateurC(r) + deltaGroupe
                    End If
                Next
            End If
        Next

        ' 7. Écrire les parts dans le DataTable
        '    Note : depParCreateurA/C contiennent déjà les parts individuelles correctes
        '    En mode inégalitaire avec plusieurs membres de même catégorie :
        '    redistribuer le total catégorie (hors bonus inédit) selon PH
        For Each row As DataRow In lignesA
            EcrirePartsRow(row, depParCreateurA(row), drParCreateurA(row), dt)
        Next
        For Each row As DataRow In lignesC
            EcrirePartsRow(row, depParCreateurC(row), drParCreateurC(row), dt)
        Next

        ' Éditeurs : égalitaire ou inégalitaire (par PH) selon params.Inegalitaire
        For Each lettrage As String In lettragesEditeurs
            Dim edsDuGroupe = lignesE.Where(Function(r) r("Lettrage").ToString().Trim() = lettrage).ToList()
            Dim depGroupe As Double = depEditeurParGroupe(lettrage)
            Dim drGroupe  As Double = drEditeurParGroupe(lettrage)
            If Not params.Inegalitaire OrElse edsDuGroupe.Count = 1 Then
                ' Égalitaire : parts égales entre les éditeurs du groupe
                Dim depIndiv As Double = depGroupe / edsDuGroupe.Count
                Dim drIndiv  As Double = drGroupe  / edsDuGroupe.Count
                For Each row As DataRow In edsDuGroupe
                    EcrirePartsRow(row, depIndiv, drIndiv, dt)
                Next
            Else
                ' Inégalitaire : pondération par PH dans le groupe
                Dim sommePHGroupe As Double = edsDuGroupe.Sum(Function(r) ParsePH(r))
                For Each row As DataRow In edsDuGroupe
                    Dim ph As Double = ParsePH(row)
                    Dim w As Double = If(sommePHGroupe > 0, ph / sommePHGroupe, 1.0 / edsDuGroupe.Count)
                    EcrirePartsRow(row, depGroupe * w, drGroupe * w, dt)
                Next
            End If
        Next

        ' Arrangeurs / Adaptateurs : égalitaire
        RepartirCategorie(lignesAR, depAR, drAR, params.Inegalitaire, dt)
        RepartirCategorie(lignesAD, depAD, drAD, params.Inegalitaire, dt)

        ' 8. Domaine public : déléguer à l'ancienne logique
        If params.EstDomainePublic Then
            Dim depA2 As Double = 0, depC2 As Double = 0, depE2 As Double = 0
            Dim depAR2 As Double = 0, depAD2 As Double = 0
            Dim drA2  As Double = 0, drC2  As Double = 0, drE2  As Double = 0
            Dim drAR2 As Double = 0, drAD2 As Double = 0
            Dim sommePHEditeurs As Double = lignesE.Sum(Function(r) ParsePH(r))
            CalculerPartsDEP(params, nbA, nbC, nbE, aArrangeur, aAdaptateur,
                             depA2, depC2, depE2, depAR2, depAD2)
            CalculerPartsDR(params, nbA, nbC, nbE, aArrangeur, aAdaptateur,
                            sommePHEditeurs, drA2, drC2, drE2, drAR2, drAD2)
            RepartirCategorie(lignesA,  depA2, drA2, params.Inegalitaire, dt)
            RepartirCategorie(lignesC,  depC2, drC2, params.Inegalitaire, dt)
            Dim sommePHE2 As Double = lignesE.Sum(Function(r) ParsePH(r))
            RepartirEditeurs(lignesE, depE2, drE2, sommePHE2, params.Inegalitaire, dt)
            RepartirCategorie(lignesAR, depAR2, drAR2, params.Inegalitaire, dt)
            RepartirCategorie(lignesAD, depAD2, drAD2, params.Inegalitaire, dt)
        End If

        ' 9. Arrondir et ajuster totaux à 100
        AjusterTotal100(dt, "PH")
        AjusterTotal100(dt, "DE")
        AjusterTotal100(dt, "DR")
    End Sub

    ''' Arrondit tous les valeurs d'une colonne a 3 decimales
    ''' puis ajuste la derniere ligne pour que le total soit exactement 100.
    Private Shared Sub AjusterTotal100(dt As DataTable, colName As String)
        If Not dt.Columns.Contains(colName) Then Return
        Dim rows As New List(Of DataRow)(dt.Rows.Cast(Of DataRow)())
        If rows.Count = 0 Then Return
        Dim total As Double = 0
        For Each row As DataRow In rows
            Dim v As Double = 0
            Double.TryParse(row(colName).ToString().Replace(",", "."),
                            Globalization.NumberStyles.Any,
                            Globalization.CultureInfo.InvariantCulture, v)
            Dim rounded As Double = Math.Round(v, 3)
            row(colName) = rounded.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
            total += rounded
        Next
        Dim ecart As Double = Math.Round(100 - total, 3)
        If ecart <> 0 Then
            Dim lastRow As DataRow = rows(rows.Count - 1)
            Dim lastVal As Double = 0
            Double.TryParse(lastRow(colName).ToString().Replace(",", "."),
                            Globalization.NumberStyles.Any,
                            Globalization.CultureInfo.InvariantCulture, lastVal)
            Dim corrected As Double = Math.Round(lastVal + ecart, 3)
            lastRow(colName) = corrected.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
        End If
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
    ''' En inégalitaire : pondération par PH dans le groupe du même lettrage.
    ''' Chaque groupe de lettrage reçoit une part proportionnelle à son PH total
    ''' puis la redistribue en interne selon les PH individuels.
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
            ' Inégalitaire — pondération par lettrage
            ' Étape 1 : part de chaque groupe de lettrage = PH total groupe / PH total E
            Dim totalPHE As Double = lignesE.Sum(Function(r) ParsePH(r))
            If totalPHE = 0 Then
                RepartirEditeurs(lignesE, partGlobaleDEP, partGlobaleDR, 0, False, dt)
                Return
            End If

            Dim groupes = lignesE.GroupBy(Function(r) r("Lettrage").ToString()).ToList()
            For Each groupe In groupes
                Dim edsDuGroupe = groupe.ToList()
                Dim phGroupe As Double = edsDuGroupe.Sum(Function(r) ParsePH(r))

                ' Part DEP/DR du groupe proportionnelle à son PH total
                Dim depGroupe As Double = partGlobaleDEP * phGroupe / totalPHE
                Dim drGroupe  As Double = partGlobaleDR  * phGroupe / totalPHE

                If phGroupe = 0 OrElse edsDuGroupe.Count = 1 Then
                    ' Égalitaire dans le groupe
                    Dim depIndiv As Double = Math.Round(depGroupe / edsDuGroupe.Count, 4)
                    Dim drIndiv  As Double = Math.Round(drGroupe  / edsDuGroupe.Count, 4)
                    For Each row As DataRow In edsDuGroupe
                        EcrirePartsRow(row, depIndiv, drIndiv, dt)
                    Next
                Else
                    ' Inégalitaire dans le groupe — pondération par PH individuel
                    For Each row As DataRow In edsDuGroupe
                        Dim ph As Double = ParsePH(row)
                        Dim depIndiv As Double = Math.Round(depGroupe * ph / phGroupe, 4)
                        Dim drIndiv  As Double = Math.Round(drGroupe  * ph / phGroupe, 4)
                        EcrirePartsRow(row, depIndiv, drIndiv, dt)
                    Next
                End If
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
    ''' Recalcule les PH de tous les ayants droit après ajout/suppression.
    ''' Règle : somme PH = 100% toujours.
    ''' PH créateur = part DR théorique de sa catégorie - somme PH de ses éditeurs.
    ''' PH éditeurs = valeurs contractuelles conservées (ou initialisées égalitaires).
    ''' </summary>
    ''' <summary>
    ''' Calcule les PH par défaut : même logique que DR.
    ''' A=25%, C=25%, E=50% réparti égalitairement.
    ''' Chaque catégorie est indépendante.
    ''' </summary>
    Public Shared Sub RecalculerPHApresAjout(dt As DataTable, params As ParamsOeuvre)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim lignesA  = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "A").ToList()
        Dim lignesC  = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "C").ToList()
        Dim lignesE  = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "E").ToList()
        Dim lignesAR = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "AR").ToList()
        Dim lignesAD = dt.AsEnumerable().Where(Function(r) r("Role").ToString() = "AD").ToList()

        Dim aA  As Boolean = lignesA.Count  > 0
        Dim aC  As Boolean = lignesC.Count  > 0
        Dim aE  As Boolean = lignesE.Count  > 0
        Dim aAR As Boolean = lignesAR.Count > 0
        Dim aAD As Boolean = lignesAD.Count > 0

        ' Parts de référence (même logique que DR par défaut)
        Dim drA As Double = 0, drC As Double = 0, drE As Double = 0
        Dim drAR As Double = 0, drAD As Double = 0
        CalculerPartsDR(params, lignesA.Count, lignesC.Count, lignesE.Count,
                        aAR, aAD, 50.0, drA, drC, drE, drAR, drAD)

        ' Parts inédites de référence (sans éditeurs → 100% pour créateurs)
        Dim drAInedit As Double = 0, drCInedit As Double = 0, drEInedit As Double = 0
        Dim drARInedit As Double = 0, drADInedit As Double = 0
        CalculerPartsDR(params, lignesA.Count, lignesC.Count, 0,
                        aAR, aAD, 0.0, drAInedit, drCInedit, drEInedit, drARInedit, drADInedit)

        ' Part inédite individuelle = part totale du groupe (créateur + ses éditeurs)
        Dim partGroupeA  As Double = If(lignesA.Count  > 0, drAInedit  / lignesA.Count,  0)
        Dim partGroupeC  As Double = If(lignesC.Count  > 0, drCInedit  / lignesC.Count,  0)
        Dim partGroupeAR As Double = If(lignesAR.Count > 0, drARInedit / lignesAR.Count, 0)
        Dim partGroupeAD As Double = If(lignesAD.Count > 0, drADInedit / lignesAD.Count, 0)

        ' ── Étape 1 : PH créateurs sans éditeurs → part inédite complète ────────
        Dim lettragesAvecEditeurs As New HashSet(Of String)(
            lignesE.Select(Function(r) r("Lettrage").ToString()))

        For Each r As DataRow In lignesA
            If Not lettragesAvecEditeurs.Contains(r("Lettrage").ToString()) Then
                r("PH") = partGroupeA.ToString(Globalization.CultureInfo.InvariantCulture)
            End If
        Next
        For Each r As DataRow In lignesC
            If Not lettragesAvecEditeurs.Contains(r("Lettrage").ToString()) Then
                r("PH") = partGroupeC.ToString(Globalization.CultureInfo.InvariantCulture)
            End If
        Next

        ' ── Étape 2 : PH créateurs AVEC éditeurs → partGroupe / (1 + sommePartsE) ──
        ' On calcule d'abord le PH créateur, PUIS on calcule les éditeurs dessus
        If aE Then
            Dim groupesE = lignesE.GroupBy(Function(r) r("Lettrage").ToString()).ToList()
            For Each groupe In groupesE
                Dim lettrage As String = groupe.Key
                Dim edsDuGroupe = groupe.ToList()

                Dim crea = dt.AsEnumerable().FirstOrDefault(
                    Function(r) (r("Role").ToString() = "A" OrElse r("Role").ToString() = "C") AndAlso
                                 r("Lettrage").ToString() = lettrage)
                Dim partGroupe As Double = If(crea Is Nothing, 0,
                    If(crea("Role").ToString() = "A", partGroupeA, partGroupeC))

                ' Parts relatives des éditeurs (normalisées sur 1.0)
                Dim sommePHGroupe As Double = edsDuGroupe.Sum(Function(r) ParsePH(r))
                Dim partsRelatives As New Dictionary(Of DataRow, Double)()
                If sommePHGroupe = 0 Then
                    Dim partEgale As Double = 1.0 / edsDuGroupe.Count
                    For Each r As DataRow In edsDuGroupe
                        partsRelatives(r) = partEgale
                    Next
                Else
                    For Each r As DataRow In edsDuGroupe
                        partsRelatives(r) = ParsePH(r) / sommePHGroupe
                    Next
                End If

                ' totalUnites = 1 (créateur) + 1 (éditeurs ensemble = 1 unité)
                Dim phCrea As Double = Math.Round(partGroupe / 2.0, 4)

                ' Mettre à jour le créateur EN PREMIER
                If crea IsNot Nothing Then
                    crea("PH") = phCrea.ToString(Globalization.CultureInfo.InvariantCulture)
                End If

                ' PH éditeur = PH réel du créateur × part relative de l'éditeur
                For Each r As DataRow In edsDuGroupe
                    r("PH") = Math.Round(phCrea * partsRelatives(r), 4).
                              ToString(Globalization.CultureInfo.InvariantCulture)
                Next
            Next
        End If

        ' AR et AD
        For Each r As DataRow In lignesAR
            r("PH") = partGroupeAR.ToString(Globalization.CultureInfo.InvariantCulture)
        Next
        For Each r As DataRow In lignesAD
            r("PH") = partGroupeAD.ToString(Globalization.CultureInfo.InvariantCulture)
        Next

        ' AR : répartition égale
        If aAR Then
            Dim phAR As Double = Math.Round(drAR / lignesAR.Count, 4)
            For Each r As DataRow In lignesAR
                r("PH") = phAR.ToString(Globalization.CultureInfo.InvariantCulture)
            Next
        End If

        ' AD : répartition égale
        If aAD Then
            Dim phAD As Double = Math.Round(drAD / lignesAD.Count, 4)
            For Each r As DataRow In lignesAD
                r("PH") = phAD.ToString(Globalization.CultureInfo.InvariantCulture)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Rééquilibre après modification manuelle d'un PH.
    ''' - A ou C modifié → ajuste les autres A/C ET met à jour ses éditeurs (même lettrage)
    ''' - E modifié → ajuste les autres E du même lettrage
    ''' totalAvant = total de la sous-catégorie AVANT modification.
    ''' </summary>
    Public Shared Sub RééquilibrerCategorie(dt As DataTable, rowModifiee As DataRow, totalAvant As Double)
        Dim role As String = rowModifiee("Role").ToString().Trim().ToUpper()
        Dim lettrage As String = rowModifiee("Lettrage").ToString().Trim()
        Dim nouvPH As Double = ParsePH(rowModifiee)

        ' ── Cas créateur (A ou C) ───────────────────────────────────────────────
        If role = "A" OrElse role = "C" Then

            ' 1. Ajuster les autres créateurs de la même sous-catégorie
            Dim autres = dt.AsEnumerable().
                Where(Function(r) r("Role").ToString() = role AndAlso Not r Is rowModifiee).ToList()

            If autres.Count > 0 Then
                Dim reste As Double = totalAvant - nouvPH
                Dim sommePHAutres As Double = autres.Sum(Function(r) ParsePH(r))
                If sommePHAutres = 0 OrElse reste <= 0 Then
                    Dim phIndiv As Double = If(reste > 0, Math.Round(reste / autres.Count, 4), 0)
                    For Each r As DataRow In autres
                        r("PH") = Math.Max(0, phIndiv).ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                Else
                    For Each r As DataRow In autres
                        Dim ph As Double = ParsePH(r)
                        r("PH") = Math.Max(0, Math.Round(reste * ph / sommePHAutres, 4)).
                                  ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                End If
            End If

            ' 2. Mettre à jour les éditeurs de TOUS les créateurs touchés
            '    (le créateur modifié ET les autres ajustés)
            Dim tousLesCreas = New List(Of DataRow)(autres) From {rowModifiee}
            For Each crea As DataRow In tousLesCreas
                Dim lettrCrea As String = crea("Lettrage").ToString().Trim()
                Dim phCrea As Double = ParsePH(crea)
                Dim editeursDuCrea = dt.AsEnumerable().
                    Where(Function(r) r("Role").ToString() = "E" AndAlso
                                      r("Lettrage").ToString() = lettrCrea).ToList()
                If editeursDuCrea.Count = 0 Then Continue For

                Dim sommePHEd As Double = editeursDuCrea.Sum(Function(r) ParsePH(r))
                If sommePHEd = 0 Then
                    Dim phIndiv As Double = Math.Round(phCrea / editeursDuCrea.Count, 4)
                    For Each r As DataRow In editeursDuCrea
                        r("PH") = phIndiv.ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                Else
                    For Each r As DataRow In editeursDuCrea
                        Dim ph As Double = ParsePH(r)
                        r("PH") = Math.Round(phCrea * ph / sommePHEd, 4).
                                  ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                End If
            Next

        ' ── Cas éditeur (E) ─────────────────────────────────────────────────────
        ElseIf role = "E" Then
            Dim autres = dt.AsEnumerable().
                Where(Function(r) r("Role").ToString() = "E" AndAlso
                                  r("Lettrage").ToString() = lettrage AndAlso
                                  Not r Is rowModifiee).ToList()

            If autres.Count > 0 Then
                Dim reste As Double = totalAvant - nouvPH
                Dim sommePHAutres As Double = autres.Sum(Function(r) ParsePH(r))
                If sommePHAutres = 0 OrElse reste <= 0 Then
                    Dim phIndiv As Double = If(reste > 0, Math.Round(reste / autres.Count, 4), 0)
                    For Each r As DataRow In autres
                        r("PH") = Math.Max(0, phIndiv).ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                Else
                    For Each r As DataRow In autres
                        Dim ph As Double = ParsePH(r)
                        r("PH") = Math.Max(0, Math.Round(reste * ph / sommePHAutres, 4)).
                                  ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                End If
            End If

        ' ── Cas AR / AD ─────────────────────────────────────────────────────────
        ElseIf role = "AR" OrElse role = "AD" Then
            Dim autres = dt.AsEnumerable().
                Where(Function(r) r("Role").ToString() = role AndAlso Not r Is rowModifiee).ToList()
            If autres.Count > 0 Then
                Dim reste As Double = totalAvant - nouvPH
                Dim sommePHAutres As Double = autres.Sum(Function(r) ParsePH(r))
                If sommePHAutres = 0 OrElse reste <= 0 Then
                    Dim phIndiv As Double = If(reste > 0, Math.Round(reste / autres.Count, 4), 0)
                    For Each r As DataRow In autres
                        r("PH") = Math.Max(0, phIndiv).ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                Else
                    For Each r As DataRow In autres
                        Dim ph As Double = ParsePH(r)
                        r("PH") = Math.Max(0, Math.Round(reste * ph / sommePHAutres, 4)).
                                  ToString(Globalization.CultureInfo.InvariantCulture)
                    Next
                End If
            End If
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────
    ' UTILITAIRES
    ' ─────────────────────────────────────────────────────────────

    Private Shared Sub EcrirePartsRow(row As DataRow, dep As Double, dr As Double, dt As DataTable)
        If dt.Columns.Contains("DE") Then
            row("DE") = Math.Round(dep, 3).ToString("0.###", Globalization.CultureInfo.InvariantCulture)
        End If
        If dt.Columns.Contains("DR") Then
            row("DR") = Math.Round(dr, 3).ToString("0.###", Globalization.CultureInfo.InvariantCulture)
        End If
    End Sub

    Public Shared Function IsEditeur(role As String) As Boolean
        Return role.Trim().ToUpper() = "E"
    End Function

    Public Shared Function ParsePH(row As DataRow) As Double
        Dim ph As Double = 0
        Double.TryParse(row("PH").ToString().Replace(",", "."),
                        Globalization.NumberStyles.Any,
                        Globalization.CultureInfo.InvariantCulture, ph)
        Return ph
    End Function

End Class
