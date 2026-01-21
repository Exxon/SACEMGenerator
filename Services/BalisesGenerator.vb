Imports System.Text
Imports System.Globalization

''' <summary>
''' Générateur de balises pour templates SACEM
''' Reproduit la logique Python : balises simples, indexées et calculées
''' </summary>
Public Class BalisesGenerator
    Private ReadOnly _data As SACEMData
    Private _balises As Dictionary(Of String, String)

    Public Sub New(data As SACEMData)
        _data = data
        _balises = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Sub

    ''' <summary>
    ''' Génère toutes les balises (simples + indexées + calculées)
    ''' </summary>
    Public Function GenerateAllBalises() As Dictionary(Of String, String)
        _balises.Clear()

        ' 1. Balises simples (champs directs)
        GenerateSimpleBalises()

        ' 2. Balises de dates décomposées
        GenerateDateBalises()

        ' 3. Balises de durée décomposées
        GenerateDureeBalises()

        ' 4. Balises indexées pour BDO (Role1, Designation1, etc.)
        GenerateIndexedBalises()

        ' 5. Balises calculées (listes, pourcentages, etc.)
        GenerateCalculatedBalises()

        ' 6. Balises de territoire et arrangement
        GenerateConditionalBalises()

        Return _balises
    End Function

    ''' <summary>
    ''' Génère les balises simples (champs directs)
    ''' </summary>
    Private Sub GenerateSimpleBalises()
        SetBalise("Titre", _data.Titre)
        SetBalise("SousTitre", _data.SousTitre)
        SetBalise("Interprete", _data.Interprete)
        SetBalise("Duree", _data.Duree)
        SetBalise("Genre", _data.Genre)
        SetBalise("Date", _data.Date)
        SetBalise("ISWC", _data.ISWC)
        SetBalise("Lieu", _data.Lieu)
        SetBalise("Territoire", _data.Territoire)
        SetBalise("Arrangement", _data.Arrangement)
        SetBalise("Commentaire", _data.Commentaire)
        SetBalise("Faita", _data.Faita)
        SetBalise("Faitle", _data.Faitle)
        SetBalise("Declaration", _data.Declaration)
        SetBalise("Format", _data.Format)
    End Sub

    ''' <summary>
    ''' Génère les balises de dates décomposées
    ''' </summary>
    Private Sub GenerateDateBalises()
        ' Date principale
        Dim dateInfo As DateInfo = DateInfo.Parse(_data.Date)
        SetBalise("JJ", dateInfo.JJ)
        SetBalise("MM", dateInfo.MM)
        SetBalise("AAAA", dateInfo.AAAA)

        ' Date "Fait le"
        Dim faitLeInfo As DateInfo = DateInfo.Parse(_data.Faitle)
        SetBalise("FJJ", faitLeInfo.JJ)
        SetBalise("FMM", faitLeInfo.MM)
        SetBalise("FAAAA", faitLeInfo.AAAA)
    End Sub

    ''' <summary>
    ''' Génère les balises de durée décomposées
    ''' </summary>
    Private Sub GenerateDureeBalises()
        Dim dureeInfo As DureeInfo = DureeInfo.Parse(_data.Duree)
        SetBalise("hh", dureeInfo.HH)
        SetBalise("mm", dureeInfo.MM)
        SetBalise("ss", dureeInfo.SS)
    End Sub

    ''' <summary>
    ''' Génère les balises indexées pour le BDO
    ''' Role1, Designation1, Lettrage1, COAD/IPI1, PH1, etc.
    ''' </summary>
    Private Sub GenerateIndexedBalises()
        For i As Integer = 0 To _data.AyantsDroit.Count - 1
            Dim ayant As AyantDroit = _data.AyantsDroit(i)
            Dim index As Integer = i + 1

            SetBalise($"Role{index}", ayant.BDO.Role)
            SetBalise($"Designation{index}", ayant.Identite.Designation)
            SetBalise($"Lettrage{index}", ayant.BDO.Lettrage)
            SetBalise($"COAD/IPI{index}", ayant.BDO.COAD_IPI)
            SetBalise($"PH{index}", ayant.BDO.PH)
            
            ' Date contrat placeholder
            SetBalise($"DateContrat{index}", $"#dteAyt{index}#")
        Next
    End Sub

    ''' <summary>
    ''' Génère les balises calculées (reproduction logique Python)
    ''' </summary>
    Private Sub GenerateCalculatedBalises()
        ' Liste des compositeurs
        Dim compositeurs As New List(Of String)
        Dim compositeursPseudo As New List(Of String)
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "C" OrElse ayant.BDO.Role = "AR" Then
                ' Pseudonyme si rempli, sinon Prénom Nom
                Dim displayName As String
                If Not String.IsNullOrEmpty(ayant.Identite.Pseudonyme) Then
                    displayName = ayant.Identite.Pseudonyme
                Else
                    Dim prenom As String = If(ayant.Identite.Prenom, "")
                    Dim nom As String = If(ayant.Identite.Nom, "")
                    displayName = $"{prenom} {nom}".Trim()
                    If String.IsNullOrEmpty(displayName) Then
                        displayName = ayant.Identite.Designation
                    End If
                End If
                
                If Not compositeurs.Contains(displayName) Then
                    compositeurs.Add(displayName)
                End If
                
                ' Pour pseudolist, garder la même logique
                Dim pseudo As String = If(Not String.IsNullOrEmpty(ayant.Identite.Pseudonyme), 
                                          ayant.Identite.Pseudonyme, 
                                          ayant.Identite.Designation)
                If Not compositeursPseudo.Contains(pseudo) Then
                    compositeursPseudo.Add(pseudo)
                End If
            End If
        Next
        SetBalise("compositeurslist", FormatList(compositeurs))
        SetBalise("compositeurspseudolist", FormatList(compositeursPseudo))

        ' Liste des auteurs
        Dim auteurs As New List(Of String)
        Dim auteursPseudo As New List(Of String)
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "A" OrElse ayant.BDO.Role = "AD" Then
                ' Pseudonyme si rempli, sinon Prénom Nom
                Dim displayName As String
                If Not String.IsNullOrEmpty(ayant.Identite.Pseudonyme) Then
                    displayName = ayant.Identite.Pseudonyme
                Else
                    Dim prenom As String = If(ayant.Identite.Prenom, "")
                    Dim nom As String = If(ayant.Identite.Nom, "")
                    displayName = $"{prenom} {nom}".Trim()
                    If String.IsNullOrEmpty(displayName) Then
                        displayName = ayant.Identite.Designation
                    End If
                End If
                
                If Not auteurs.Contains(displayName) Then
                    auteurs.Add(displayName)
                End If
                
                ' Pour pseudolist, garder la même logique
                Dim pseudo As String = If(Not String.IsNullOrEmpty(ayant.Identite.Pseudonyme), 
                                          ayant.Identite.Pseudonyme, 
                                          ayant.Identite.Designation)
                If Not auteursPseudo.Contains(pseudo) Then
                    auteursPseudo.Add(pseudo)
                End If
            End If
        Next
        SetBalise("auteurslist", FormatList(auteurs))
        SetBalise("auteurspseudolist", FormatList(auteursPseudo))

        ' Liste des éditeurs (tous)
        Dim editeurs As New List(Of String)
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "E" Then
                If Not editeurs.Contains(ayant.Identite.Designation) Then
                    editeurs.Add(ayant.Identite.Designation)
                End If
            End If
        Next
        SetBalise("editeurslist", FormatList(editeurs))
        SetBalise("editeurslistou", FormatListOu(editeurs))
        SetBalise("editeurslistetou", FormatListEtOu(editeurs))
        SetBalise("editeurslistoude", FormatListOuDe(editeurs))

        ' Éditeurs sans format
        Dim editeursSansFormat As New List(Of String)
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "E" AndAlso ayant.Identite.Designation <> _data.Format Then
                If Not editeursSansFormat.Contains(ayant.Identite.Designation) Then
                    editeursSansFormat.Add(ayant.Identite.Designation)
                End If
            End If
        Next
        SetBalise("editnoformat", FormatList(editeursSansFormat))
        SetBalise("editnoformatoua", FormatListOuA(editeursSansFormat))

        ' Liste des sub-éditeurs formatée
        ' Format: EDITEUR (pour son propre compte) ou EDITEUR (pour son propre compte et pour le compte de X, et pour le compte de Y)
        SetBalise("sublist", GenerateSubList())

        ' Crédits
        SetBalise("credits", String.Join(" / ", editeurs))

        ' Répartition des parts éditeurs (editsplit)
        GenerateEditSplit()
    End Sub
    
    ''' <summary>
    ''' Génère la liste des éditeurs avec leur rôle (propre compte / pour le compte de...)
    ''' </summary>
    Private Function GenerateSubList() As String
        ' Collecter tous les éditeurs uniques
        Dim editeursUniques As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role <> "E" Then Continue For
            
            Dim designation As String = GetDesignation(ayant)
            If String.IsNullOrEmpty(designation) Then Continue For
            
            If Not editeursUniques.ContainsKey(designation.ToUpper()) Then
                editeursUniques(designation.ToUpper()) = designation
            End If
        Next
        
        ' Identifier qui gère qui (Managelic = l'éditeur qui gère)
        ' Clé = éditeur principal, Valeur = liste des éditeurs qu'il gère
        Dim gestionnaires As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        Dim editeursGeres As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role <> "E" Then Continue For
            
            Dim managelic As String = If(ayant.BDO.Managelic, "").Trim()
            If String.IsNullOrEmpty(managelic) Then Continue For
            
            Dim designation As String = GetDesignation(ayant)
            If String.IsNullOrEmpty(designation) Then Continue For
            
            ' Cet éditeur est géré par managelic
            editeursGeres.Add(designation.ToUpper())
            
            ' Ajouter à la liste des éditeurs gérés par managelic
            If Not gestionnaires.ContainsKey(managelic.ToUpper()) Then
                gestionnaires(managelic.ToUpper()) = New List(Of String)
            End If
            
            If Not gestionnaires(managelic.ToUpper()).Contains(designation) Then
                gestionnaires(managelic.ToUpper()).Add(designation)
            End If
        Next
        
        ' Construire la liste finale (seulement les éditeurs non gérés par d'autres)
        Dim resultats As New List(Of String)
        
        For Each kvp In editeursUniques
            Dim designation As String = kvp.Value
            Dim key As String = kvp.Key
            
            ' Ignorer les éditeurs qui sont gérés par d'autres
            If editeursGeres.Contains(key) Then Continue For
            
            ' Construire la chaîne pour cet éditeur
            If gestionnaires.ContainsKey(key) AndAlso gestionnaires(key).Count > 0 Then
                ' Cet éditeur gère d'autres éditeurs
                Dim geres As List(Of String) = gestionnaires(key)
                Dim pourLeCompte As String = String.Join(", et pour le compte de ", geres)
                resultats.Add($"{designation} (pour son propre compte et pour le compte de {pourLeCompte})")
            Else
                ' Éditeur autonome
                resultats.Add($"{designation} (pour son propre compte)")
            End If
        Next
        
        ' Formater avec virgules et "et" pour le dernier
        If resultats.Count = 0 Then
            Return ""
        ElseIf resultats.Count = 1 Then
            Return resultats(0)
        Else
            Dim allButLast As String = String.Join(", ", resultats.Take(resultats.Count - 1))
            Return $"{allButLast} et {resultats.Last()}"
        End If
    End Function
    
    ''' <summary>
    ''' Obtient la désignation d'un ayant droit (Designation pour Moral, Nom Prénom pour Physique)
    ''' </summary>
    Private Function GetDesignation(ayant As AyantDroit) As String
        If ayant.Identite.Type?.ToLower() = "moral" Then
            Return If(ayant.Identite.Designation, "").Trim()
        Else
            ' Physique : Nom Prénom
            Dim nom As String = If(ayant.Identite.Nom, "").Trim()
            Dim prenom As String = If(ayant.Identite.Prenom, "").Trim()
            Return $"{nom} {prenom}".Trim()
        End If
    End Function

    ''' <summary>
    ''' Génère la répartition des parts éditeurs (editsplit)
    ''' </summary>
    Private Sub GenerateEditSplit()
        Dim pourcentagesParEditeur As New Dictionary(Of String, Double)

        ' Calculer la somme des pourcentages par éditeur
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "E" Then
                Dim ph As Double
                If Double.TryParse(ayant.BDO.PH, NumberStyles.Any, CultureInfo.InvariantCulture, ph) Then
                    If pourcentagesParEditeur.ContainsKey(ayant.Identite.Designation) Then
                        pourcentagesParEditeur(ayant.Identite.Designation) += ph
                    Else
                        pourcentagesParEditeur(ayant.Identite.Designation) = ph
                    End If
                End If
            End If
        Next

        ' Générer les lignes editsplit
        Dim editsplit As New List(Of String)
        For Each kvp In pourcentagesParEditeur.OrderBy(Function(x) x.Key)
            Dim phText As String = NombreEnLettres(kvp.Value)
            editsplit.Add($"{kvp.Value:F2}% ({phText} pour cent) : {kvp.Key}")
        Next

        SetBalise("editsplit", String.Join(vbCrLf, editsplit))
    End Sub

    ''' <summary>
    ''' Génère les balises conditionnelles (X pour les checkboxes)
    ''' </summary>
    Private Sub GenerateConditionalBalises()
        ' Territoire
        SetBalise("M", If(_data.Territoire = "Monde", "X", ""))
        SetBalise("A", If(_data.Territoire <> "Monde", "X", ""))
        
        ' Inégalitaire
        SetBalise("P", If(_data.Inegalitaire = "TRUE", "X", ""))
        
        ' Arrangement
        SetBalise("T", If(_data.Arrangement = "Toutes", "X", ""))
        SetBalise("D", If(Not String.IsNullOrEmpty(_data.Arrangement) AndAlso _data.Arrangement <> "Toutes", "X", ""))
    End Sub

    ''' <summary>
    ''' Formate une liste : "A, B et C"
    ''' </summary>
    Private Function FormatList(items As List(Of String)) As String
        If items.Count = 0 Then Return ""
        If items.Count = 1 Then Return items(0)
        Return String.Join(", ", items.Take(items.Count - 1)) & " et " & items.Last()
    End Function

    ''' <summary>
    ''' Formate une liste avec "ou" : "A ou B ou C"
    ''' </summary>
    Private Function FormatListOu(items As List(Of String)) As String
        Return String.Join(" ou ", items)
    End Function

    ''' <summary>
    ''' Formate une liste avec "et/ou" : "A et/ou B et/ou C"
    ''' </summary>
    Private Function FormatListEtOu(items As List(Of String)) As String
        Return String.Join(" et/ou ", items)
    End Function

    ''' <summary>
    ''' Formate une liste avec "ou de" : "A ou de B ou de C"
    ''' </summary>
    Private Function FormatListOuDe(items As List(Of String)) As String
        Return String.Join(" ou de ", items)
    End Function

    ''' <summary>
    ''' Formate une liste avec "ou à" : "A ou à B ou à C"
    ''' </summary>
    Private Function FormatListOuA(items As List(Of String)) As String
        Return String.Join(" ou à ", items)
    End Function

    ''' <summary>
    ''' Définit une balise
    ''' </summary>
    Private Sub SetBalise(key As String, value As String)
        If String.IsNullOrEmpty(value) Then value = ""
        _balises(key) = value
    End Sub

    ''' <summary>
    ''' Convertit un nombre en lettres (français)
    ''' Reproduction simple de num2words
    ''' </summary>
    Private Function NombreEnLettres(nombre As Double) As String
        Try
            ' Version simplifiée pour les pourcentages
            Dim partieEntiere As Integer = CInt(Math.Floor(nombre))
            Dim partieDecimale As Integer = CInt((nombre - partieEntiere) * 100)

            Dim result As String = ConvertirNombreEnLettres(partieEntiere)

            If partieDecimale > 0 Then
                result &= " virgule " & ConvertirNombreEnLettres(partieDecimale)
            End If

            Return result

        Catch
            Return nombre.ToString("F2")
        End Try
    End Function
    
    ''' <summary>
    ''' Convertit un nombre entier (0-99) en lettres françaises
    ''' </summary>
    Private Function ConvertirNombreEnLettres(nombre As Integer) As String
        Dim units() As String = {"zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf"}
        Dim teens() As String = {"dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"}
        Dim tens() As String = {"", "", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante-dix", "quatre-vingt", "quatre-vingt-dix"}

        If nombre = 0 Then
            Return "zéro"
        ElseIf nombre < 10 Then
            Return units(nombre)
        ElseIf nombre < 20 Then
            Return teens(nombre - 10)
        ElseIf nombre < 100 Then
            Dim dizaine As Integer = nombre \ 10
            Dim unite As Integer = nombre Mod 10
            
            ' Gestion spéciale pour 70-79 et 90-99 (soixante-dix, quatre-vingt-dix)
            If dizaine = 7 Then
                ' 70-79 : soixante-dix, soixante-et-onze, soixante-douze...
                If unite = 0 Then
                    Return "soixante-dix"
                ElseIf unite = 1 Then
                    Return "soixante-et-onze"
                Else
                    Return "soixante-" & teens(unite)
                End If
            ElseIf dizaine = 9 Then
                ' 90-99 : quatre-vingt-dix, quatre-vingt-onze...
                If unite = 0 Then
                    Return "quatre-vingt-dix"
                Else
                    Return "quatre-vingt-" & teens(unite)
                End If
            ElseIf dizaine = 8 Then
                ' 80-89 : quatre-vingts, quatre-vingt-un...
                If unite = 0 Then
                    Return "quatre-vingts"
                Else
                    Return "quatre-vingt-" & units(unite)
                End If
            Else
                ' 20-69 : vingt, vingt-et-un, vingt-deux...
                Dim result As String = tens(dizaine)
                If unite = 1 Then
                    result &= "-et-un"
                ElseIf unite > 0 Then
                    result &= "-" & units(unite)
                End If
                Return result
            End If
        Else
            Return nombre.ToString()
        End If
    End Function

    ''' <summary>
    ''' Récupère les balises générées
    ''' </summary>
    Public ReadOnly Property Balises As Dictionary(Of String, String)
        Get
            Return _balises
        End Get
    End Property
End Class
