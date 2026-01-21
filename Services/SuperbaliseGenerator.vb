Imports System.Text
Imports System.Linq

''' <summary>
''' Générateur de superbalises (contenu dynamique complexe)
''' Reproduit les fonctions Python : auteurspart_main, editeurspart_main, etc.
''' Supporte le formatage Gras/Italique/Souligné
''' </summary>
Public Class SuperbaliseGenerator
    Private ReadOnly _data As SACEMData
    Private ReadOnly _paragraphReader As ParagraphTemplateReader

    Public Sub New(data As SACEMData, paragraphReader As ParagraphTemplateReader)
        _data = data
        _paragraphReader = paragraphReader
    End Sub

    ''' <summary>
    ''' Génère le contenu formaté pour {auteurspart}
    ''' Retourne une liste de segments avec formatage
    ''' </summary>
    Public Function GenerateAuteursPartFormatted() As List(Of FormattedSegment)
        Try
            Dim allSegments As New List(Of FormattedSegment)
            Dim isFirst As Boolean = True

            ' Créer une liste combinée avec gestion des rôles A+C → AC
            Dim combinedAyants As Dictionary(Of String, AyantDroit) = CombineRolesAC()

            ' Traiter chaque ayant droit (sauf éditeurs)
            For Each kvp In combinedAyants
                Dim ayant As AyantDroit = kvp.Value
                
                ' Exclure les éditeurs
                If ayant.BDO.Role = "E" Then Continue For

                ' Déterminer le template à utiliser
                Dim templateKey As String = GetTemplateKey(ayant)
                
                If String.IsNullOrEmpty(templateKey) Then
                    Debug.WriteLine($"Aucun template trouvé pour {ayant.Identite.Designation}")
                    Continue For
                End If

                ' Ajouter séparateur si pas le premier
                If Not isFirst Then
                    allSegments.Add(New FormattedSegment(vbCrLf & vbCrLf & "Et," & vbCrLf & vbCrLf))
                End If
                isFirst = False

                ' Appliquer le template avec formatage
                Dim segments As List(Of FormattedSegment) = _paragraphReader.ApplyTemplateFormatted(templateKey, ayant)
                allSegments.AddRange(segments)
            Next

            Return allSegments

        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateAuteursPartFormatted: {ex.Message}")
            Return New List(Of FormattedSegment)
        End Try
    End Function

    ''' <summary>
    ''' Génère le contenu pour {auteurspart} (version texte simple pour compatibilité)
    ''' </summary>
    Public Function GenerateAuteursPart() As String
        Try
            Dim resultats As New List(Of String)
            Dim combinedAyants As Dictionary(Of String, AyantDroit) = CombineRolesAC()

            For Each kvp In combinedAyants
                Dim ayant As AyantDroit = kvp.Value
                If ayant.BDO.Role = "E" Then Continue For

                Dim templateKey As String = GetTemplateKey(ayant)
                If String.IsNullOrEmpty(templateKey) Then Continue For

                Dim paragraphe As String = _paragraphReader.ApplyTemplate(templateKey, ayant)
                If Not String.IsNullOrEmpty(paragraphe) Then
                    resultats.Add(paragraphe)
                End If
            Next

            resultats = resultats.Distinct().OrderBy(Function(x) x).ToList()
            Dim separateur As String = vbCrLf & vbCrLf & "Et," & vbCrLf & vbCrLf
            Return String.Join(separateur, resultats)

        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateAuteursPart: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Génère le contenu formaté pour {editeurspart}
    ''' SANS DOUBLONS : un éditeur n'apparaît qu'une seule fois même s'il est dans plusieurs lettrages
    ''' </summary>
    Public Function GenerateEditeursPartFormatted() As List(Of FormattedSegment)
        Try
            Dim allSegments As New List(Of FormattedSegment)
            Dim isFirst As Boolean = True
            
            ' Dédupliquer les éditeurs par Designation (clé normalisée)
            Dim editeursUniques As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                If ayant.BDO.Role <> "E" Then Continue For
                
                Dim key As String = NormalizeDesignation(ayant.Identite.Designation)
                If Not editeursUniques.ContainsKey(key) Then
                    editeursUniques(key) = ayant
                End If
            Next

            ' Générer les paragraphes pour chaque éditeur unique
            For Each kvp In editeursUniques
                Dim ayant As AyantDroit = kvp.Value
                
                Dim templateKey As String = GetTemplateKey(ayant)
                If String.IsNullOrEmpty(templateKey) Then Continue For

                If Not isFirst Then
                    allSegments.Add(New FormattedSegment(vbCrLf & vbCrLf & "Et," & vbCrLf & vbCrLf))
                End If
                isFirst = False

                Dim segments As List(Of FormattedSegment) = _paragraphReader.ApplyTemplateFormatted(templateKey, ayant)
                allSegments.AddRange(segments)
            Next

            Return allSegments

        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateEditeursPartFormatted: {ex.Message}")
            Return New List(Of FormattedSegment)
        End Try
    End Function

    ''' <summary>
    ''' Génère le contenu pour {editeurspart} (version texte simple)
    ''' SANS DOUBLONS
    ''' </summary>
    Public Function GenerateEditeursPart() As String
        Try
            Dim resultats As New List(Of String)
            
            ' Dédupliquer les éditeurs
            Dim editeursUniques As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                If ayant.BDO.Role <> "E" Then Continue For
                
                Dim key As String = NormalizeDesignation(ayant.Identite.Designation)
                If Not editeursUniques.ContainsKey(key) Then
                    editeursUniques(key) = ayant
                End If
            Next

            For Each kvp In editeursUniques
                Dim ayant As AyantDroit = kvp.Value
                
                Dim templateKey As String = GetTemplateKey(ayant)
                If String.IsNullOrEmpty(templateKey) Then Continue For

                Dim paragraphe As String = _paragraphReader.ApplyTemplate(templateKey, ayant)
                If Not String.IsNullOrEmpty(paragraphe) Then
                    resultats.Add(paragraphe)
                End If
            Next

            Dim separateur As String = vbCrLf & vbCrLf & "Et," & vbCrLf & vbCrLf
            Return String.Join(separateur, resultats)

        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateEditeursPart: {ex.Message}")
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Normalise une designation pour la comparaison (majuscules, sans espaces multiples)
    ''' </summary>
    Private Function NormalizeDesignation(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        
        Dim result As String = text.ToUpper().Trim()
        
        While result.Contains("  ")
            result = result.Replace("  ", " ")
        End While
        
        Return result
    End Function

    ''' <summary>
    ''' Génère le contenu pour {subpart}
    ''' Le nombre d'itérations = le nombre d'entrées dans [sublist]
    ''' Groupé par éditeur (ou groupe d'éditeurs si Managesub)
    ''' [createur] = tous les A/C des lettrages où cet éditeur apparaît
    ''' [editeurcede] = éditeurs qui cèdent leur sous-édition à cet éditeur (via Managesub)
    ''' </summary>
    Public Function GenerateSubPart() As String
        Try
            Dim resultats As New List(Of String)
            
            ' Étape 1 : Identifier les éditeurs principaux (ceux qui n'ont pas de Managesub)
            ' et ceux qui cèdent leur sous-édition (ceux qui ont un Managesub)
            Dim editeursPrincipaux As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
            Dim editeursQuiCedent As New Dictionary(Of String, List(Of AyantDroit))(StringComparer.OrdinalIgnoreCase)
            Dim lettragesParEditeur As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                If ayant.BDO.Role <> "E" Then Continue For
                
                Dim designation As String = GetDesignationForDisplay(ayant)
                If String.IsNullOrEmpty(designation) Then Continue For
                
                Dim key As String = designation.ToUpper()
                Dim managesub As String = If(ayant.BDO.Managesub, "").Trim()
                Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
                
                If String.IsNullOrEmpty(managesub) Then
                    ' Éditeur principal (gère sa propre sous-édition)
                    If Not editeursPrincipaux.ContainsKey(key) Then
                        editeursPrincipaux(key) = ayant
                        lettragesParEditeur(key) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    End If
                    If Not String.IsNullOrEmpty(lettrage) Then
                        lettragesParEditeur(key).Add(lettrage)
                    End If
                Else
                    ' Cet éditeur cède sa sous-édition à managesub
                    Dim principalKey As String = managesub.ToUpper()
                    If Not editeursQuiCedent.ContainsKey(principalKey) Then
                        editeursQuiCedent(principalKey) = New List(Of AyantDroit)
                    End If
                    ' Éviter les doublons
                    If Not editeursQuiCedent(principalKey).Any(Function(e) GetDesignationForDisplay(e).ToUpper() = key) Then
                        editeursQuiCedent(principalKey).Add(ayant)
                    End If
                    ' Ajouter aussi cet éditeur comme principal pour son propre lettrage
                    If Not editeursPrincipaux.ContainsKey(key) Then
                        editeursPrincipaux(key) = ayant
                        lettragesParEditeur(key) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    End If
                    If Not String.IsNullOrEmpty(lettrage) Then
                        lettragesParEditeur(key).Add(lettrage)
                    End If
                End If
            Next
            
            ' Étape 2 : Pour chaque éditeur principal, trouver tous les créateurs de ses lettrages
            For Each kvp In editeursPrincipaux
                Dim editeurKey As String = kvp.Key
                Dim editeur As AyantDroit = kvp.Value
                Dim editeurDesignation As String = GetDesignationForDisplay(editeur)
                
                ' Vérifier si cet éditeur cède sa sous-édition à un autre (alors on ne génère pas pour lui)
                Dim managesub As String = If(editeur.BDO.Managesub, "").Trim()
                If Not String.IsNullOrEmpty(managesub) Then
                    ' Cet éditeur cède sa sous-édition, on ne génère pas de ligne pour lui
                    Continue For
                End If
                
                ' Récupérer tous les lettrages de cet éditeur
                Dim lettrages As HashSet(Of String) = If(lettragesParEditeur.ContainsKey(editeurKey), lettragesParEditeur(editeurKey), New HashSet(Of String))
                
                ' Ajouter aussi les lettrages des éditeurs qui cèdent à celui-ci
                If editeursQuiCedent.ContainsKey(editeurKey) Then
                    For Each editeurCedant In editeursQuiCedent(editeurKey)
                        Dim cedantKey As String = GetDesignationForDisplay(editeurCedant).ToUpper()
                        If lettragesParEditeur.ContainsKey(cedantKey) Then
                            For Each l In lettragesParEditeur(cedantKey)
                                lettrages.Add(l)
                            Next
                        End If
                    Next
                End If
                
                ' Trouver tous les créateurs (A/C) dans ces lettrages
                Dim createursUniques As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
                
                ' Vérifier si l'éditeur actuel est un EACA (il est aussi créateur dans ses propres lettrages)
                Dim editeurEstEACA As Boolean = False
                Dim editeurNom As String = ""
                Dim editeurPrenom As String = ""
                
                If editeur.Identite.Type?.ToLower() <> "moral" Then
                    editeurNom = If(editeur.Identite.Nom, "").Trim().ToUpper()
                    editeurPrenom = If(editeur.Identite.Prenom, "").Trim().ToUpper()
                End If
                
                For Each ayant In _data.AyantsDroit
                    If ayant.BDO.Role <> "A" AndAlso ayant.BDO.Role <> "C" Then Continue For
                    
                    Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
                    If String.IsNullOrEmpty(lettrage) OrElse Not lettrages.Contains(lettrage) Then Continue For
                    
                    Dim createurKey As String = GetPersonKey(ayant)
                    
                    ' Vérifier si ce créateur est l'éditeur actuel (cas EACA)
                    Dim estLEditeurActuel As Boolean = False
                    If ayant.Identite.Type?.ToLower() <> "moral" AndAlso Not String.IsNullOrEmpty(editeurNom) Then
                        Dim nomCreateur As String = If(ayant.Identite.Nom, "").Trim().ToUpper()
                        Dim prenomCreateur As String = If(ayant.Identite.Prenom, "").Trim().ToUpper()
                        If nomCreateur = editeurNom AndAlso prenomCreateur = editeurPrenom Then
                            estLEditeurActuel = True
                            editeurEstEACA = True
                        End If
                    End If
                    
                    ' Ajouter à la liste des créateurs (sauf si c'est l'éditeur lui-même)
                    If Not estLEditeurActuel Then
                        If Not createursUniques.ContainsKey(createurKey) Then
                            createursUniques(createurKey) = ayant
                        End If
                    End If
                Next
                
                ' Récupérer les éditeurs qui cèdent leur sous-édition à cet éditeur
                Dim editeursCedants As List(Of AyantDroit) = Nothing
                If editeursQuiCedent.ContainsKey(editeurKey) Then
                    editeursCedants = editeursQuiCedent(editeurKey)
                End If
                
                ' Construire [editeurcede] si des éditeurs cèdent leur sous-édition
                Dim editeurcedeTexte As String = ""
                If editeursCedants IsNot Nothing AndAlso editeursCedants.Count > 0 Then
                    Dim cedantsNoms As New List(Of String)
                    For Each cedant In editeursCedants
                        Dim cedantNom As String = GetDesignationForDisplay(cedant)
                        If Not cedantsNoms.Contains(cedantNom) Then
                            cedantsNoms.Add(cedantNom)
                        End If
                    Next
                    editeurcedeTexte = FormatListeEt(cedantsNoms)
                End If
                
                ' Générer le texte
                If createursUniques.Count > 0 Then
                    ' Template SUBS - l'éditeur règle aux créateurs
                    Dim createursList As New List(Of String)
                    For Each createur In createursUniques.Values
                        createursList.Add(GetCreateurNomComplet(createur))
                    Next
                    Dim createursTexte As String = FormatListeEt(createursList)
                    
                    Dim template As String = _paragraphReader.GetTemplate("SUBS")
                    If Not String.IsNullOrEmpty(template) Then
                        template = template.Replace("[editeur]", editeurDesignation)
                        template = template.Replace("[createur]", createursTexte)
                        
                        ' Gérer [editeurcedesub] et le bloc associé
                        If String.IsNullOrEmpty(editeurcedeTexte) Then
                            ' Pas d'éditeurs qui cèdent : supprimer le bloc ", et ceux de [editeurcedesub],"
                            template = template.Replace(", et ceux de [editeurcedesub],", ",")
                            template = template.Replace(" et ceux de [editeurcedesub]", "")
                            template = template.Replace("[editeurcedesub]", "")
                        Else
                            ' Avec éditeurs qui cèdent
                            template = template.Replace("[editeurcedesub]", editeurcedeTexte)
                            
                            ' Si plusieurs éditeurs cèdent OU plusieurs créateurs, lui → leur
                            If editeurcedeTexte.Contains(" et ") OrElse createursTexte.Contains(" et ") Then
                                template = template.Replace("lui revenant", "leur revenant")
                                template = template.Replace("la part des redevances lui", "la part des redevances leur")
                            End If
                        End If
                        
                        resultats.Add(template.Trim())
                    End If
                End If
                
                ' Si l'éditeur est aussi EACA (créateur dans ses propres lettrages), générer ligne SUBEAC
                If editeurEstEACA Then
                    Dim template As String = _paragraphReader.GetTemplate("SUBEAC")
                    If Not String.IsNullOrEmpty(template) Then
                        template = template.Replace("[createur]", editeurDesignation)
                        template = template.Replace("[designation]", editeurDesignation)
                        resultats.Add(template.Trim())
                    End If
                End If
            Next
            
            ' Joindre avec double saut de ligne pour séparer chaque paragraphe
            Return String.Join(vbCrLf & vbCrLf, resultats)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateSubPart: {ex.Message}")
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Applique le template START_SUBS pour un éditeur et un créateur
    ''' Template: Sur les sommes perçues par lui, [editeur] se chargera de régler à [createur]...
    ''' </summary>
    Private Function ApplyTemplateSubs(editeur As AyantDroit, createur As AyantDroit) As String
        Try
            Dim template As String = _paragraphReader.GetTemplate("SUBS")
            If String.IsNullOrEmpty(template) Then Return ""
            
            ' Remplacer les balises éditeur
            Dim editeurDesignation As String = GetDesignationForDisplay(editeur)
            template = template.Replace("[editeur]", editeurDesignation)
            template = template.Replace("[designation]", editeurDesignation)
            
            ' Remplacer les balises créateur
            Dim createurNom As String = GetCreateurNomComplet(createur)
            template = template.Replace("[createur]", createurNom)
            template = template.Replace("[createurnom]", createurNom)
            
            Return template
        Catch ex As Exception
            Debug.WriteLine($"Erreur ApplyTemplateSubs: {ex.Message}")
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Applique le template START_SUBS2 pour un éditeur principal, plusieurs créateurs et sous-éditeurs
    ''' Template: Sur les sommes perçues par lui, [editeur] se chargera de régler à [createur]...
    ''' [editeur] = éditeur principal
    ''' [createur] = liste des créateurs (A/C)
    ''' [editeurcede] = liste des sous-éditeurs (ceux gérés par l'éditeur principal)
    ''' </summary>
    Private Function ApplyTemplateSubs2(editeurPrincipal As AyantDroit, createurs As List(Of AyantDroit), sousEditeurs As List(Of AyantDroit)) As String
        Try
            Dim template As String = _paragraphReader.GetTemplate("SUBS2")
            If String.IsNullOrEmpty(template) Then Return ""
            
            ' Éditeur principal
            Dim editeurDesignation As String = GetDesignationForDisplay(editeurPrincipal)
            template = template.Replace("[editeur]", editeurDesignation)
            template = template.Replace("[designation]", editeurDesignation)
            
            ' Liste des créateurs (Nom Prénom)
            Dim createursNoms As New List(Of String)
            For Each createur In createurs
                createursNoms.Add(GetCreateurNomComplet(createur))
            Next
            Dim createursTexte As String = FormatListeEt(createursNoms)
            template = template.Replace("[createur]", createursTexte)
            template = template.Replace("[createurs]", createursTexte)
            
            ' Liste des sous-éditeurs (editeurcede)
            Dim sousEditeursNoms As New List(Of String)
            For Each sousEd In sousEditeurs
                Dim sousEdNom As String = GetDesignationForDisplay(sousEd)
                If Not sousEditeursNoms.Contains(sousEdNom) Then
                    sousEditeursNoms.Add(sousEdNom)
                End If
            Next
            Dim editeurcedeTexte As String = FormatListeEt(sousEditeursNoms)
            template = template.Replace("[editeurcede]", editeurcedeTexte)
            template = template.Replace("[sousediteurs]", editeurcedeTexte)
            template = template.Replace("[subediteurs]", editeurcedeTexte)
            
            Return template
        Catch ex As Exception
            Debug.WriteLine($"Erreur ApplyTemplateSubs2: {ex.Message}")
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Obtient la clé unique pour identifier une personne (pour détecter EACA)
    ''' </summary>
    Private Function GetPersonKey(ayant As AyantDroit) As String
        If ayant.Identite.Type?.ToLower() = "moral" Then
            Return If(ayant.Identite.Designation, "").Trim().ToUpper()
        Else
            ' Physique : Nom + Prénom normalisés
            Dim nom As String = If(ayant.Identite.Nom, "").Trim().ToUpper()
            Dim prenom As String = If(ayant.Identite.Prenom, "").Trim().ToUpper()
            Return $"{nom}|{prenom}"
        End If
    End Function
    
    ''' <summary>
    ''' Obtient la désignation pour affichage (Moral: Designation, Physique: Nom Prénom)
    ''' </summary>
    Private Function GetDesignationForDisplay(ayant As AyantDroit) As String
        If ayant.Identite.Type?.ToLower() = "moral" Then
            Return If(ayant.Identite.Designation, "").Trim()
        Else
            Dim nom As String = If(ayant.Identite.Nom, "").Trim()
            Dim prenom As String = If(ayant.Identite.Prenom, "").Trim()
            Return $"{nom} {prenom}".Trim()
        End If
    End Function
    
    ''' <summary>
    ''' Obtient le nom complet du créateur : Nom Prénom (sans pseudo pour subpart)
    ''' </summary>
    Private Function GetCreateurNomComplet(ayant As AyantDroit) As String
        Dim nom As String = If(ayant.Identite.Nom, "").Trim()
        Dim prenom As String = If(ayant.Identite.Prenom, "").Trim()
        
        Return $"{nom} {prenom}".Trim()
    End Function
    
    ''' <summary>
    ''' Formate une liste avec virgules et "et" pour le dernier
    ''' </summary>
    Private Function FormatListeEt(items As List(Of String)) As String
        If items.Count = 0 Then Return ""
        If items.Count = 1 Then Return items(0)
        
        Dim allButLast As String = String.Join(", ", items.Take(items.Count - 1))
        Return $"{allButLast} et {items.Last()}"
    End Function

    ''' <summary>
    ''' Génère le contenu pour {licpart}
    ''' Même logique que {subpart} mais basée sur Managelic
    ''' [editeur] = éditeur principal
    ''' [createur] = créateurs des lettrages
    ''' [editeurcedelic] = éditeurs qui cèdent leur gestion de licences (via Managelic)
    ''' </summary>
    Public Function GenerateLicPart() As String
        Try
            Dim resultats As New List(Of String)
            
            ' Étape 1 : Identifier les éditeurs principaux (ceux qui n'ont pas de Managelic)
            ' et ceux qui cèdent leur gestion de licences (ceux qui ont un Managelic)
            Dim editeursPrincipaux As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
            Dim editeursQuiCedent As New Dictionary(Of String, List(Of AyantDroit))(StringComparer.OrdinalIgnoreCase)
            Dim lettragesParEditeur As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                If ayant.BDO.Role <> "E" Then Continue For
                
                Dim designation As String = GetDesignationForDisplay(ayant)
                If String.IsNullOrEmpty(designation) Then Continue For
                
                Dim key As String = designation.ToUpper()
                Dim managelic As String = If(ayant.BDO.Managelic, "").Trim()
                Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
                
                If String.IsNullOrEmpty(managelic) Then
                    ' Éditeur principal (gère ses propres licences)
                    If Not editeursPrincipaux.ContainsKey(key) Then
                        editeursPrincipaux(key) = ayant
                        lettragesParEditeur(key) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    End If
                    If Not String.IsNullOrEmpty(lettrage) Then
                        lettragesParEditeur(key).Add(lettrage)
                    End If
                Else
                    ' Cet éditeur cède sa gestion de licences à managelic
                    Dim principalKey As String = managelic.ToUpper()
                    If Not editeursQuiCedent.ContainsKey(principalKey) Then
                        editeursQuiCedent(principalKey) = New List(Of AyantDroit)
                    End If
                    ' Éviter les doublons
                    If Not editeursQuiCedent(principalKey).Any(Function(e) GetDesignationForDisplay(e).ToUpper() = key) Then
                        editeursQuiCedent(principalKey).Add(ayant)
                    End If
                    ' Ajouter aussi cet éditeur comme principal pour son propre lettrage
                    If Not editeursPrincipaux.ContainsKey(key) Then
                        editeursPrincipaux(key) = ayant
                        lettragesParEditeur(key) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    End If
                    If Not String.IsNullOrEmpty(lettrage) Then
                        lettragesParEditeur(key).Add(lettrage)
                    End If
                End If
            Next
            
            ' Étape 2 : Pour chaque éditeur principal, trouver tous les créateurs de ses lettrages
            For Each kvp In editeursPrincipaux
                Dim editeurKey As String = kvp.Key
                Dim editeur As AyantDroit = kvp.Value
                Dim editeurDesignation As String = GetDesignationForDisplay(editeur)
                
                ' Vérifier si cet éditeur cède sa gestion de licences à un autre (alors on ne génère pas pour lui)
                Dim managelic As String = If(editeur.BDO.Managelic, "").Trim()
                If Not String.IsNullOrEmpty(managelic) Then
                    ' Cet éditeur cède sa gestion, on ne génère pas de ligne pour lui
                    Continue For
                End If
                
                ' Récupérer tous les lettrages de cet éditeur
                Dim lettrages As HashSet(Of String) = If(lettragesParEditeur.ContainsKey(editeurKey), lettragesParEditeur(editeurKey), New HashSet(Of String))
                
                ' Ajouter aussi les lettrages des éditeurs qui cèdent à celui-ci
                If editeursQuiCedent.ContainsKey(editeurKey) Then
                    For Each editeurCedant In editeursQuiCedent(editeurKey)
                        Dim cedantKey As String = GetDesignationForDisplay(editeurCedant).ToUpper()
                        If lettragesParEditeur.ContainsKey(cedantKey) Then
                            For Each l In lettragesParEditeur(cedantKey)
                                lettrages.Add(l)
                            Next
                        End If
                    Next
                End If
                
                ' Trouver tous les créateurs (A/C) dans ces lettrages
                Dim createursUniques As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
                
                ' Vérifier si l'éditeur actuel est un EACA (il est aussi créateur dans ses propres lettrages)
                Dim editeurEstEACA As Boolean = False
                Dim editeurNom As String = ""
                Dim editeurPrenom As String = ""
                
                If editeur.Identite.Type?.ToLower() <> "moral" Then
                    editeurNom = If(editeur.Identite.Nom, "").Trim().ToUpper()
                    editeurPrenom = If(editeur.Identite.Prenom, "").Trim().ToUpper()
                End If
                
                For Each ayant In _data.AyantsDroit
                    If ayant.BDO.Role <> "A" AndAlso ayant.BDO.Role <> "C" Then Continue For
                    
                    Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
                    If String.IsNullOrEmpty(lettrage) OrElse Not lettrages.Contains(lettrage) Then Continue For
                    
                    Dim createurKey As String = GetPersonKey(ayant)
                    
                    ' Vérifier si ce créateur est l'éditeur actuel (cas EACA)
                    Dim estLEditeurActuel As Boolean = False
                    If ayant.Identite.Type?.ToLower() <> "moral" AndAlso Not String.IsNullOrEmpty(editeurNom) Then
                        Dim nomCreateur As String = If(ayant.Identite.Nom, "").Trim().ToUpper()
                        Dim prenomCreateur As String = If(ayant.Identite.Prenom, "").Trim().ToUpper()
                        If nomCreateur = editeurNom AndAlso prenomCreateur = editeurPrenom Then
                            estLEditeurActuel = True
                            editeurEstEACA = True
                        End If
                    End If
                    
                    ' Ajouter à la liste des créateurs (sauf si c'est l'éditeur lui-même)
                    If Not estLEditeurActuel Then
                        If Not createursUniques.ContainsKey(createurKey) Then
                            createursUniques(createurKey) = ayant
                        End If
                    End If
                Next
                
                ' Récupérer les éditeurs qui cèdent leur gestion de licences à cet éditeur
                Dim editeursCedants As List(Of AyantDroit) = Nothing
                If editeursQuiCedent.ContainsKey(editeurKey) Then
                    editeursCedants = editeursQuiCedent(editeurKey)
                End If
                
                ' Construire [editeurcedelic] si des éditeurs cèdent leur gestion de licences
                Dim editeurcedeTexte As String = ""
                If editeursCedants IsNot Nothing AndAlso editeursCedants.Count > 0 Then
                    Dim cedantsNoms As New List(Of String)
                    For Each cedant In editeursCedants
                        Dim cedantNom As String = GetDesignationForDisplay(cedant)
                        If Not cedantsNoms.Contains(cedantNom) Then
                            cedantsNoms.Add(cedantNom)
                        End If
                    Next
                    editeurcedeTexte = FormatListeEt(cedantsNoms)
                End If
                
                ' Générer le texte
                If createursUniques.Count > 0 Then
                    ' Template LIC
                    Dim createursList As New List(Of String)
                    For Each createur In createursUniques.Values
                        createursList.Add(GetCreateurNomComplet(createur))
                    Next
                    Dim createursTexte As String = FormatListeEt(createursList)
                    
                    Dim template As String = _paragraphReader.GetTemplate("LIC")
                    If Not String.IsNullOrEmpty(template) Then
                        template = template.Replace("[editeur]", editeurDesignation)
                        template = template.Replace("[createur]", createursTexte)
                        
                        ' Gérer [editeurcedelic] et le bloc associé
                        If String.IsNullOrEmpty(editeurcedeTexte) Then
                            ' Pas d'éditeurs qui cèdent : supprimer le bloc ", et ceux de [editeurcedelic],"
                            template = template.Replace(", et ceux de [editeurcedelic],", ",")
                            template = template.Replace(" et ceux de [editeurcedelic]", "")
                            template = template.Replace("[editeurcedelic]", "")
                        Else
                            ' Avec éditeurs qui cèdent
                            template = template.Replace("[editeurcedelic]", editeurcedeTexte)
                            
                            ' Si plusieurs éditeurs cèdent OU plusieurs créateurs, lui → leur
                            If editeurcedeTexte.Contains(" et ") OrElse createursTexte.Contains(" et ") Then
                                template = template.Replace("lui revenant", "leur revenant")
                                template = template.Replace("la part lui", "la part leur")
                            End If
                        End If
                        
                        resultats.Add(template.Trim())
                    End If
                End If
                
                ' Si l'éditeur est aussi EACA (créateur dans ses propres lettrages), générer ligne LICEAC
                If editeurEstEACA Then
                    Dim template As String = _paragraphReader.GetTemplate("LICEAC")
                    If Not String.IsNullOrEmpty(template) Then
                        template = template.Replace("[createur]", editeurDesignation)
                        template = template.Replace("[designation]", editeurDesignation)
                        resultats.Add(template.Trim())
                    End If
                End If
            Next
            
            ' Joindre avec double saut de ligne pour séparer chaque paragraphe
            Return String.Join(vbCrLf & vbCrLf, resultats)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateLicPart: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Combine les rôles A et C en AC pour un même ayant droit
    ''' Utilise une clé normalisée pour détecter les doublons
    ''' </summary>
    Private Function CombineRolesAC() As Dictionary(Of String, AyantDroit)
        Dim combined As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)

        For Each ayant In _data.AyantsDroit
            ' Créer une clé normalisée
            Dim key As String = NormalizeDesignation(ayant.Identite.Designation)
            Dim role As String = ayant.BDO.Role

            If combined.ContainsKey(key) Then
                ' Combiner A+C en AC
                If (role = "A" OrElse role = "C") AndAlso
                   (combined(key).BDO.Role = "A" OrElse combined(key).BDO.Role = "C") Then
                    combined(key).BDO.Role = "AC"
                End If
            Else
                Dim copie As New AyantDroit()
                CopyAyantDroit(ayant, copie)
                combined(key) = copie
            End If
        Next

        Return combined
    End Function

    ''' <summary>
    ''' Détermine la clé du template à utiliser selon le type d'ayant droit
    ''' </summary>
    Private Function GetTemplateKey(ayant As AyantDroit) As String
        If ayant.Identite.Type = "Physique" Then
            Return "Physique"
        Else
            Dim forme As String = If(ayant.Identite.FormeJuridique, "").ToUpper().Trim()

            If String.IsNullOrEmpty(forme) Then
                Debug.WriteLine($"Forme juridique manquante pour {ayant.Identite.Designation} - utilisation du template SA par défaut")
                Return "SA"
            End If

            Select Case forme
                Case "SA", "SAS", "SASU", "SARL", "SNC", "SCI", "GMBH", "LLC", "LTD", "INC", "CORP"
                    Return "SA"
                Case "EURL"
                    Return "EURL"
                Case "EI", "AUTO-ENTREPRENEUR", "MICRO", "MICRO-ENTREPRENEUR"
                    Return "EI"
                Case "ASS", "ASSOCIATION", "ASSO"
                    Return "ASS"
                Case Else
                    Debug.WriteLine($"Forme juridique non reconnue: '{forme}' pour {ayant.Identite.Designation} - utilisation du template SA")
                    Return "SA"
            End Select
        End If
    End Function

    ''' <summary>
    ''' Copie un ayant droit
    ''' </summary>
    Private Sub CopyAyantDroit(source As AyantDroit, dest As AyantDroit)
        dest.Identite.Designation = source.Identite.Designation
        dest.Identite.Type = source.Identite.Type
        dest.Identite.Pseudonyme = source.Identite.Pseudonyme
        dest.Identite.Nom = source.Identite.Nom
        dest.Identite.Prenom = source.Identite.Prenom
        dest.Identite.Genre = source.Identite.Genre
        dest.Identite.FormeJuridique = source.Identite.FormeJuridique
        dest.Identite.Capital = source.Identite.Capital
        dest.Identite.RCS = source.Identite.RCS
        dest.Identite.Siren = source.Identite.Siren
        dest.Identite.PrenomRepresentant = source.Identite.PrenomRepresentant
        dest.Identite.NomRepresentant = source.Identite.NomRepresentant
        dest.Identite.GenreRepresentant = source.Identite.GenreRepresentant
        dest.Identite.FonctionRepresentant = source.Identite.FonctionRepresentant
        dest.Identite.Nele = source.Identite.Nele
        dest.Identite.Nea = source.Identite.Nea

        dest.BDO.Role = source.BDO.Role
        dest.BDO.COAD_IPI = source.BDO.COAD_IPI
        dest.BDO.PH = source.BDO.PH
        dest.BDO.Lettrage = source.BDO.Lettrage
        dest.BDO.Managelic = source.BDO.Managelic
        dest.BDO.Managesub = source.BDO.Managesub

        dest.Adresse.NumVoie = source.Adresse.NumVoie
        dest.Adresse.TypeVoie = source.Adresse.TypeVoie
        dest.Adresse.NomVoie = source.Adresse.NomVoie
        dest.Adresse.CP = source.Adresse.CP
        dest.Adresse.Ville = source.Adresse.Ville
        dest.Adresse.Pays = source.Adresse.Pays

        dest.Contact.Mail = source.Contact.Mail
        dest.Contact.Tel = source.Contact.Tel
    End Sub

    ''' <summary>
    ''' Construit la table de gestion (Managesub ou Managelic)
    ''' </summary>
    Private Function BuildTabManage(manageType As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)

        For Each ayant In _data.AyantsDroit
            Dim manageValue As String = ""
            If manageType = "Managesub" Then
                manageValue = ayant.BDO.Managesub
            ElseIf manageType = "Managelic" Then
                manageValue = ayant.BDO.Managelic
            End If

            If Not String.IsNullOrEmpty(manageValue) Then
                Dim createur As String = ayant.Identite.Designation
                If result.ContainsKey(createur) Then
                    result(createur) &= ";" & manageValue
                Else
                    result(createur) = manageValue
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Crée un ayant droit simple pour les templates
    ''' </summary>
    Private Function CreateSimpleAyant(designation As String) As AyantDroit
        Dim ayant As New AyantDroit()
        ayant.Identite.Designation = designation
        Return ayant
    End Function
End Class
