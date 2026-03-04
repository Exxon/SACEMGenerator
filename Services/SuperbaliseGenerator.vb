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
            ' FILTRE NON-SACEM : seuls les membres SACEM sont inclus
            Dim editeursUniques As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                If ayant.BDO.Role <> "E" Then Continue For
                If Not IsSACEMMember(ayant) Then Continue For ' Exclure NON-SACEM
                If Not IsSignataire(ayant) Then Continue For  ' Exclure non-signataires
                
                Dim key As String = NormalizeDesignation(ayant.Identite.Designation)
                If Not editeursUniques.ContainsKey(key) Then
                    editeursUniques(key) = ayant
                End If
            Next

            ' Calculer le total PH SACEM éditeurs pour recalculer sur 100%
            Dim totalSACEM As Double = GetTotalSACEMPH({"E"})

            ' Générer les paragraphes pour chaque éditeur unique
            For Each kvp In editeursUniques
                Dim ayant As AyantDroit = kvp.Value

                ' Recalculer le PH sur 100% de la part éditeur SACEM
                Dim phBrut As Double
                Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, phBrut)
                Dim copie As New AyantDroit()
                CopyAyantDroit(ayant, copie)
                copie.BDO.PH = RecalculatePH(phBrut, totalSACEM)
                
                Dim templateKey As String = GetTemplateKey(copie)
                If String.IsNullOrEmpty(templateKey) Then Continue For

                If Not isFirst Then
                    allSegments.Add(New FormattedSegment(vbCrLf & vbCrLf & "Et," & vbCrLf & vbCrLf))
                End If
                isFirst = False

                Dim segments As List(Of FormattedSegment) = _paragraphReader.ApplyTemplateFormatted(templateKey, copie)
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
    ''' FILTRE NON-SACEM : seuls les membres SACEM sont inclus
    ''' </summary>
    Public Function GenerateEditeursPart() As String
        Try
            Dim resultats As New List(Of String)
            
            ' Dédupliquer les éditeurs
            Dim editeursUniques As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                If ayant.BDO.Role <> "E" Then Continue For
                If Not IsSACEMMember(ayant) Then Continue For ' Exclure NON-SACEM
                If Not IsSignataire(ayant) Then Continue For  ' Exclure non-signataires
                
                Dim key As String = NormalizeDesignation(ayant.Identite.Designation)
                If Not editeursUniques.ContainsKey(key) Then
                    editeursUniques(key) = ayant
                End If
            Next

            ' Calculer le total PH SACEM éditeurs pour recalculer sur 100%
            Dim totalSACEM As Double = GetTotalSACEMPH({"E"})

            For Each kvp In editeursUniques
                Dim ayant As AyantDroit = kvp.Value

                ' Recalculer le PH sur 100% de la part éditeur SACEM
                Dim phBrut As Double
                Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, phBrut)
                Dim copie As New AyantDroit()
                CopyAyantDroit(ayant, copie)
                copie.BDO.PH = RecalculatePH(phBrut, totalSACEM)
                
                Dim templateKey As String = GetTemplateKey(copie)
                If String.IsNullOrEmpty(templateKey) Then Continue For

                Dim paragraphe As String = _paragraphReader.ApplyTemplate(templateKey, copie)
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
                If Not IsSACEMMember(ayant) Then Continue For ' Exclure NON-SACEM
                If Not IsSignataire(ayant) Then Continue For  ' Exclure non-signataires
                
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
                If Not IsSACEMMember(ayant) Then Continue For ' Exclure NON-SACEM
                If Not IsSignataire(ayant) Then Continue For  ' Exclure non-signataires
                
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
    ''' Calcule le total PH SACEM pour un rôle donné ("A","C","AC" = auteurs/compositeurs, "E" = éditeurs)
    ''' et retourne un ratio pour recalculer sur 100%.
    ''' Les PH sont exprimés sur 50% (part auteur ou part éditeur de l'œuvre totale).
    ''' </summary>
    Private Function GetTotalSACEMPH(roleFilter As String()) As Double
        Dim total As Double = 0
        For Each ayant In _data.AyantsDroit
            If Not roleFilter.Contains(ayant.BDO.Role) Then Continue For
            If Not IsSACEMMember(ayant) Then Continue For
            If Not IsSignataire(ayant) Then Continue For
            Dim ph As Double
            If Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."),
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, ph) Then
                total += ph
            End If
        Next
        Return If(total > 0, total, 1) ' Éviter division par zéro
    End Function

    ''' <summary>
    ''' Recalcule un PH brut (sur 50%) en pourcentage sur 100% de la part SACEM.
    ''' </summary>
    Private Function RecalculatePH(phBrut As Double, totalSACEM As Double) As String
        Dim phRecalcule As Double = phBrut / totalSACEM * 100
        Return phRecalcule.ToString("F2", Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Function CombineRolesAC() As Dictionary(Of String, AyantDroit)
        Dim combined As New Dictionary(Of String, AyantDroit)(StringComparer.OrdinalIgnoreCase)

        ' Calculer le total PH SACEM auteurs/compositeurs pour recalculer sur 100%
        Dim totalSACEM As Double = GetTotalSACEMPH({"A", "C", "AD"})

        For Each ayant In _data.AyantsDroit
            ' Exclure les NON-SACEM
            If Not IsSACEMMember(ayant) Then Continue For
            ' Exclure les non-signataires
            If Not IsSignataire(ayant) Then Continue For
            
            ' Créer une clé normalisée basée sur Designation OU Nom+Prenom
            Dim key As String
            If Not String.IsNullOrEmpty(ayant.Identite.Designation) Then
                key = NormalizeDesignation(ayant.Identite.Designation)
            Else
                ' Pour les personnes physiques, utiliser Nom + Prenom
                key = NormalizeDesignation($"{ayant.Identite.Nom} {ayant.Identite.Prenom}")
            End If
            
            Dim role As String = ayant.BDO.Role

            If combined.ContainsKey(key) Then
                ' Combiner A+C en AC (seulement si c'est la même personne)
                If (role = "A" OrElse role = "C") AndAlso
                   (combined(key).BDO.Role = "A" OrElse combined(key).BDO.Role = "C") Then
                    combined(key).BDO.Role = "AC"
                End If
                ' Cumuler le PH brut pour recalcul
                Dim phExistant As Double
                Double.TryParse(If(combined(key).BDO.PH, "0").Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, phExistant)
                Dim phNouveau As Double
                Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, phNouveau)
                combined(key).BDO.PH = RecalculatePH(phExistant + phNouveau, totalSACEM)
            Else
                Dim copie As New AyantDroit()
                CopyAyantDroit(ayant, copie)
                ' Recalculer le PH sur 100% de la part auteur SACEM
                Dim phBrut As Double
                Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, phBrut)
                copie.BDO.PH = RecalculatePH(phBrut, totalSACEM)
                combined(key) = copie
            End If
        Next

        Return combined
    End Function
    
    ''' <summary>
    ''' Vérifie si un ayant droit est membre SACEM
    ''' Retourne True si SACEM ou si SocieteGestion est vide (défaut = SACEM)
    ''' </summary>
    Private Function IsSACEMMember(ayant As AyantDroit) As Boolean
        Dim societe As String = If(ayant.Identite.SocieteGestion, "").Trim().ToUpper()
        Return String.IsNullOrEmpty(societe) OrElse societe = "SACEM"
    End Function

    ''' <summary>
    ''' Vérifie si un ayant droit est signataire du dépôt (TRUE par défaut si absent du JSON)
    ''' </summary>
    Private Function IsSignataire(ayant As AyantDroit) As Boolean
        Return ayant.BDO.Signataire
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
        dest.BDO.Signataire = source.BDO.Signataire

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

    ' =====================================================
    ' GÉNÉRATION DES BLOCS NON-SACEM (FACTORISÉS)
    ' Blocs atomiques dans template_paragrahs.docx :
    ' - MENTION_NONSACEM : Coédition + EDITEUR
    ' - LIST_NONSACEM : Non-signataire + lettre
    ' - OGC_NONSACEM : Droits collectés OGC
    ' - BDO_NONSACEM : Commentaire BDO
    ' =====================================================

    ''' <summary>
    ''' Génère le bloc {MENTION_NONSACEM}
    ''' Coédition + "l'EDITEUR"
    ''' Utilisé dans : CCEOM Art.11, CCEOM Art.16, COED Art.3
    ''' </summary>
    Public Function GenerateMentionNonSACEM() As String
        Try
            If Not HasNonSACEM() Then Return ""
            
            Dim template As String = _paragraphReader.GetTemplate("MENTION_NONSACEM")
            If String.IsNullOrEmpty(template) Then Return ""
            
            Return ApplyNonSACEMBalises(template)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateMentionNonSACEM: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Génère le bloc {LIST_NONSACEM}
    ''' Non-signataire + "lettre de répartition ci-après annexée"
    ''' Utilisé dans : CCEOM Art.16, COED Art.3
    ''' </summary>
    Public Function GenerateListNonSACEM() As String
        Try
            If Not HasNonSACEM() Then Return ""
            
            Dim template As String = _paragraphReader.GetTemplate("LIST_NONSACEM")
            If String.IsNullOrEmpty(template) Then Return ""
            
            Return ApplyNonSACEMBalises(template)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateListNonSACEM: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Génère le bloc {OGC_NONSACEM}
    ''' "Les droits collectés par les Organismes de gestion collective..."
    ''' Utilisé dans : COED Art.3
    ''' </summary>
    Public Function GenerateOGCNonSACEM() As String
        Try
            If Not HasNonSACEM() Then Return ""
            
            Dim template As String = _paragraphReader.GetTemplate("OGC_NONSACEM")
            If String.IsNullOrEmpty(template) Then Return ""
            
            Return ApplyNonSACEMBalises(template)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateOGCNonSACEM: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Génère le bloc {BDO_NONSACEM}
    ''' Commentaire pour le BDO
    ''' </summary>
    Public Function GenerateBDONonSACEM() As String
        Try
            If Not HasNonSACEM() Then Return ""
            
            Dim template As String = _paragraphReader.GetTemplate("BDO_NONSACEM")
            If String.IsNullOrEmpty(template) Then Return ""
            
            Return ApplyNonSACEMBalises(template)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateBDONonSACEM: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Vérifie s'il y a des ayants droit NON-SACEM
    ''' </summary>
    Private Function HasNonSACEM() As Boolean
        For Each ayant In _data.AyantsDroit
            Dim societe As String = If(ayant.Identite.SocieteGestion, "").Trim().ToUpper()
            If Not String.IsNullOrEmpty(societe) AndAlso societe <> "SACEM" Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Applique les balises NON-SACEM à un template
    ''' </summary>
    Private Function ApplyNonSACEMBalises(template As String) As String
        ' Collecter les données NON-SACEM
        Dim listeAC_SACEM As New List(Of String)
        Dim listeESACEM As New List(Of String)
        Dim listeNonSACEM As New List(Of String)
        Dim listeNonSACEM_Noms As New List(Of String)
        Dim partsSACEM As Double = 0
        Dim partsNonSACEM As Double = 0
        Dim countNonSACEM As Integer = 0
        Dim editeurDeInfo As String = ""
        
        For Each ayant In _data.AyantsDroit
            Dim societe As String = If(ayant.Identite.SocieteGestion, "SACEM").Trim().ToUpper()
            Dim isSACEM As Boolean = (societe = "SACEM" OrElse String.IsNullOrEmpty(societe))
            Dim role As String = If(ayant.BDO.Role, "").Trim().ToUpper()
            Dim isAC As Boolean = (role = "A" OrElse role = "C" OrElse role = "AC" OrElse role = "AR" OrElse role = "AD")
            Dim isE As Boolean = (role = "E")
            
            Dim ph As Double = 0
            Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, ph)
            
            Dim displayName As String = GetDisplayName(ayant)
            Dim societeAffichage As String = If(ayant.Identite.SocieteGestion, "").Trim().ToUpper()
            
            If isSACEM Then
                partsSACEM += ph
                If isAC AndAlso Not listeAC_SACEM.Contains(displayName) Then
                    listeAC_SACEM.Add(displayName)
                End If
                If isE AndAlso Not listeESACEM.Contains(displayName) Then
                    listeESACEM.Add(displayName)
                End If
            Else
                partsNonSACEM += ph
                countNonSACEM += 1
                Dim nomAvecSociete As String = $"{displayName} ({societeAffichage})"
                If Not listeNonSACEM.Contains(nomAvecSociete) Then
                    listeNonSACEM.Add(nomAvecSociete)
                End If
                If Not listeNonSACEM_Noms.Contains(displayName) Then
                    listeNonSACEM_Noms.Add(displayName)
                End If
                
                ' Détecter éditeur étranger qui édite un AC SACEM
                If isE Then
                    Dim lettrage As String = If(ayant.BDO.Lettrage, "").Trim().ToUpper()
                    If Not String.IsNullOrEmpty(lettrage) Then
                        For Each autreAyant In _data.AyantsDroit
                            Dim autreSociete As String = If(autreAyant.Identite.SocieteGestion, "SACEM").Trim().ToUpper()
                            Dim autreIsSACEM As Boolean = (autreSociete = "SACEM" OrElse String.IsNullOrEmpty(autreSociete))
                            Dim autreRole As String = If(autreAyant.BDO.Role, "").Trim().ToUpper()
                            Dim autreIsAC As Boolean = (autreRole = "A" OrElse autreRole = "C" OrElse autreRole = "AC" OrElse autreRole = "AR" OrElse autreRole = "AD")
                            Dim autreLettrage As String = If(autreAyant.BDO.Lettrage, "").Trim().ToUpper()
                            
                            If autreIsSACEM AndAlso autreIsAC AndAlso autreLettrage = lettrage Then
                                editeurDeInfo = $" éditeur de {GetDisplayName(autreAyant)}"
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Next
        
        ' Formater les valeurs
        Dim strPartsSACEM As String = Math.Round(partsSACEM, 2).ToString("F2", Globalization.CultureInfo.GetCultureInfo("fr-FR")).Replace(".", ",")
        Dim strPartsNonSACEM As String = Math.Round(partsNonSACEM, 2).ToString("F2", Globalization.CultureInfo.GetCultureInfo("fr-FR")).Replace(".", ",")
        Dim strListeAC_SACEM As String = FormatListeEt(listeAC_SACEM)
        Dim strListeESACEM As String = FormatListeEt(listeESACEM)
        Dim strListeNonSACEM As String = FormatListeEt(listeNonSACEM)
        Dim strListeNonSACEM_Noms As String = FormatListeEt(listeNonSACEM_Noms)
        
        ' Pluriel/Singulier
        Dim isPluriel As Boolean = (countNonSACEM > 1)
        Dim strEstSont As String = If(isPluriel, "sont", "est")
        Dim strIlIls As String = If(isPluriel, "Ils", "Il")
        Dim strNestPas As String = If(isPluriel, "ne sont pas", "n'est pas")
        Dim strPluriel As String = If(isPluriel, "s", "")
        
        ' Remplacer les balises dans le template
        Dim result As String = template
        result = result.Replace("[ListeAC_SACEM]", strListeAC_SACEM)
        result = result.Replace("[ListeE_SACEM]", strListeESACEM)
        result = result.Replace("[ListeNonSACEM]", strListeNonSACEM)
        result = result.Replace("[ListeNonSACEM_Noms]", strListeNonSACEM_Noms)
        result = result.Replace("[PartsSACEM]", strPartsSACEM)
        result = result.Replace("[PartsNonSACEM]", strPartsNonSACEM)
        result = result.Replace("[EstSont]", strEstSont)
        result = result.Replace("[IlIls]", strIlIls)
        result = result.Replace("[NestPasNeSontPas]", strNestPas)
        result = result.Replace("[Pluriel]", strPluriel)
        result = result.Replace("[EditeurDe]", editeurDeInfo)
        
        Return result.Trim()
    End Function

    ''' <summary>
    ''' Obtient le nom d'affichage d'un ayant droit
    ''' </summary>
    Private Function GetDisplayName(ayant As AyantDroit) As String
        If ayant.Identite.Type = "Moral" Then
            Return If(ayant.Identite.Designation, "")
        End If
        
        Dim prenom As String = If(ayant.Identite.Prenom, "").Trim()
        Dim nom As String = If(ayant.Identite.Nom, "").Trim().ToUpper()
        
        If Not String.IsNullOrEmpty(prenom) AndAlso Not String.IsNullOrEmpty(nom) Then
            Return $"{prenom} {nom}"
        ElseIf Not String.IsNullOrEmpty(nom) Then
            Return nom
        ElseIf Not String.IsNullOrEmpty(prenom) Then
            Return prenom
        Else
            Return If(ayant.Identite.Designation, "")
        End If
    End Function

    ' =====================================================
    ' GENERATION DU BLOC DEPOT PARTIEL
    ' =====================================================

    Private Function HasPartiel() As Boolean
        Return _data.AyantsDroit.Any(Function(a) Not a.BDO.Signataire)
    End Function

    Public Function GenerateMentionPartiel() As String
        Try
            If Not HasPartiel() Then Return ""

            Dim template As String = _paragraphReader.GetTemplate("MENTION_PARTIEL")
            If String.IsNullOrEmpty(template) Then Return ""

            Dim totalPart As Double = 0.0
            Dim noms As New List(Of String)

            For Each ayant In _data.AyantsDroit
                If Not ayant.BDO.Signataire Then Continue For

                Dim ph As Double = 0
                Double.TryParse(If(ayant.BDO.PH, "0").Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, ph)
                totalPart += ph

                Dim nm As String = GetDisplayName(ayant)
                If Not String.IsNullOrEmpty(nm) AndAlso Not noms.Contains(nm) Then
                    noms.Add(nm)
                End If
            Next

            Dim pctStr As String = Math.Round(totalPart, 2).ToString("F2",
                Globalization.CultureInfo.GetCultureInfo("fr-FR")).Replace(".", ",")
            Dim nomsStr As String = If(noms.Any(), FormatListeEt(noms), "les signataires")

            Dim result As String = template
            result = result.Replace("[PartsSignataires]", pctStr)
            result = result.Replace("[ListeSignataires]", nomsStr)

            Return result

        Catch ex As Exception
            Debug.WriteLine($"Erreur GenerateMentionPartiel: {ex.Message}")
            Return ""
        End Try
    End Function

End Class
