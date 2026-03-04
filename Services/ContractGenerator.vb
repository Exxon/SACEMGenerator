Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Linq
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

''' <summary>
''' Générateur de contrats SACEM
''' Reproduction complète de Create_CTR_SCM.py
''' NOUVELLE VERSION : Remplacement style Find and Replace de Word
''' </summary>
Public Class ContractGenerator
    Private ReadOnly _data As SACEMData
    Private ReadOnly _paragraphReader As ParagraphTemplateReader
    Private _log As New List(Of String)

    ''' <summary>
    ''' Log des opérations
    ''' </summary>
    Public ReadOnly Property GenerationLog As List(Of String)
        Get
            Return _log
        End Get
    End Property

    ''' <summary>
    ''' Constructeur
    ''' </summary>
    Public Sub New(data As SACEMData, paragraphReader As ParagraphTemplateReader)
        _data = data
        _paragraphReader = paragraphReader
    End Sub

    ''' <summary>
    ''' Génère un contrat complet
    ''' </summary>
    Public Function Generate(templatePath As String, outputPath As String, contractType As String) As Boolean
        Try
            _log.Clear()
            _log.Add($"=== GÉNÉRATION CONTRAT {contractType} ===")
            _log.Add($"Template: {Path.GetFileName(templatePath)}")
            _log.Add($"Sortie: {Path.GetFileName(outputPath)}")

            ' Vérifications
            If Not File.Exists(templatePath) Then
                Throw New FileNotFoundException($"Template introuvable: {templatePath}")
            End If

            ' 1. Copier le template vers la destination
            File.Copy(templatePath, outputPath, True)
            _log.Add("Template copié")

            ' 2. Générer les balises calculées
            _log.Add("Génération des balises calculées...")
            Dim balisesGen As New BalisesGenerator(_data)
            Dim allBalises As Dictionary(Of String, String) = balisesGen.GenerateAllBalises()
            _log.Add($"  {allBalises.Count} balises générées")

            ' 3. Générer les superbalises
            _log.Add("Génération des superbalises...")
            Dim superbalises As Dictionary(Of String, Object) = GenerateSuperbalises(contractType)
            _log.Add($"  {superbalises.Count} superbalises générées")

            ' 4. Ouvrir le document et effectuer les remplacements
            _log.Add("Traitement du document Word...")
            Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(outputPath, True)
                Dim body As Body = wordDoc.MainDocumentPart.Document.Body

                ' ÉTAPE 1 : Fusionner les runs fragmentés pour les balises
                MergeFragmentedRuns(body)

                ' ÉTAPE 2 : Remplacer les balises simples [xxx]
                For Each kvp In allBalises
                    Dim balise As String = $"[{kvp.Key}]"
                    Dim valeur As String = If(kvp.Value, "")
                    SearchAndReplace(body, balise, valeur)
                Next

                ' ÉTAPE 3 : Remplacer les superbalises {xxx}
                For Each kvp In superbalises
                    Dim superbalise As String = $"{{{kvp.Key}}}"
                    
                    If TypeOf kvp.Value Is Table Then
                        ' C'est un tableau - traitement spécial
                        InsertTableAtPlaceholder(body, superbalise, CType(kvp.Value, Table))
                        _log.Add($"    → Tableau inséré pour {{{kvp.Key}}}")
                    ElseIf TypeOf kvp.Value Is List(Of FormattedSegment) Then
                        ' C'est une liste de segments formatés
                        Dim segments As List(Of FormattedSegment) = CType(kvp.Value, List(Of FormattedSegment))
                        InsertFormattedSegmentsAtPlaceholder(body, superbalise, segments)
                        _log.Add($"    → Contenu formaté inséré pour {{{kvp.Key}}}")
                    Else
                        ' C'est du texte simple
                        Dim valeur As String = If(kvp.Value?.ToString(), "")
                        SearchAndReplace(body, superbalise, valeur)
                        If Not String.IsNullOrEmpty(valeur) Then
                            _log.Add($"    → Texte inséré pour {{{kvp.Key}}}")
                        End If
                    End If
                Next

                ' 5. Sauvegarder
                wordDoc.MainDocumentPart.Document.Save()
            End Using

            _log.Add("✓ Document généré avec succès")
            _log.Add($"=== GÉNÉRATION TERMINÉE ===")

            Return True

        Catch ex As Exception
            _log.Add($"✗ ERREUR: {ex.Message}")
            _log.Add($"  Type: {ex.GetType().Name}")
            If ex.InnerException IsNot Nothing Then
                _log.Add($"  Détail: {ex.InnerException.Message}")
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Fusionne les runs fragmentés qui contiennent des balises
    ''' Word peut fragmenter {auteurspart} en { + auteurspart + }
    ''' Cette méthode fusionne ces runs pour permettre le remplacement
    ''' </summary>
    Private Sub MergeFragmentedRuns(body As Body)
        For Each paragraph In body.Descendants(Of Paragraph)().ToList()
            MergeRunsInParagraph(paragraph)
        Next
    End Sub

    ''' <summary>
    ''' Fusionne les runs dans un paragraphe pour reconstituer les balises fragmentées
    ''' VERSION SIMPLIFIÉE ET ROBUSTE
    ''' </summary>
    Private Sub MergeRunsInParagraph(paragraph As Paragraph)
        Dim runs = paragraph.Descendants(Of Run)().ToList()
        If runs.Count <= 1 Then Return

        ' Obtenir le texte complet du paragraphe
        Dim fullText As String = String.Join("", runs.Select(Function(r) r.InnerText))

        ' Vérifier s'il y a des balises à fusionner
        If Not (fullText.Contains("[") OrElse fullText.Contains("{")) Then Return

        ' Vérifier si une balise est fragmentée (présente dans fullText mais pas dans un seul run)
        Dim balisePattern As New Regex("\{[A-Za-z0-9_]+\}|\[[A-Za-z0-9_/]+\]")
        Dim matches = balisePattern.Matches(fullText)
        
        Dim needsMerge As Boolean = False
        For Each m As Match In matches
            Dim balise As String = m.Value
            Dim foundInSingleRun As Boolean = runs.Any(Function(r) r.InnerText.Contains(balise))
            If Not foundInSingleRun Then
                needsMerge = True
                Exit For
            End If
        Next
        
        If Not needsMerge Then Return
        
        ' APPROCHE SIMPLE: Fusionner TOUS les runs du paragraphe en un seul
        ' en préservant le formatage du premier run qui contient du texte
        Try
            ' Trouver le premier run avec du texte pour copier son formatage
            Dim firstRunWithText As Run = runs.FirstOrDefault(Function(r) Not String.IsNullOrEmpty(r.InnerText))
            If firstRunWithText Is Nothing Then Return
            
            ' Sauvegarder le formatage
            Dim savedRunProps As RunProperties = Nothing
            If firstRunWithText.RunProperties IsNot Nothing Then
                savedRunProps = CType(firstRunWithText.RunProperties.CloneNode(True), RunProperties)
            End If
            
            ' Créer un nouveau run avec tout le texte
            Dim newRun As New Run()
            If savedRunProps IsNot Nothing Then
                newRun.RunProperties = savedRunProps
            End If
            
            ' Ajouter le texte complet
            Dim newText As New Text(fullText)
            newText.Space = SpaceProcessingModeValues.Preserve
            newRun.Append(newText)
            
            ' Insérer le nouveau run avant le premier
            runs(0).InsertBeforeSelf(newRun)
            
            ' Supprimer tous les anciens runs
            For Each oldRun In runs
                oldRun.Remove()
            Next
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur fusion runs: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Recherche et remplace du texte dans tout le document (style Word Find and Replace)
    ''' </summary>
    Private Sub SearchAndReplace(body As Body, searchText As String, replaceText As String)
        For Each paragraph In body.Descendants(Of Paragraph)().ToList()
            ReplacementInParagraph(paragraph, searchText, replaceText)
        Next

        ' Chercher aussi dans les tableaux
        For Each table In body.Descendants(Of Table)().ToList()
            For Each cell In table.Descendants(Of TableCell)()
                For Each paragraph In cell.Descendants(Of Paragraph)()
                    ReplacementInParagraph(paragraph, searchText, replaceText)
                Next
            Next
        Next
    End Sub

    ''' <summary>
    ''' Remplace du texte dans un paragraphe EN PRÉSERVANT LE FORMATAGE
    ''' </summary>
    Private Sub ReplacementInParagraph(paragraph As Paragraph, searchText As String, replaceText As String)
        Try
            For Each run In paragraph.Descendants(Of Run)().ToList()
                For Each textElement In run.Descendants(Of Text)().ToList()
                    If textElement.Text.Contains(searchText) Then
                        ' Gérer les retours à la ligne dans le texte de remplacement
                        If replaceText.Contains(vbCrLf) OrElse replaceText.Contains(vbLf) Then
                            ' Remplacement multiligne - préserver le formatage du run
                            ReplaceWithMultilineText(run, textElement, searchText, replaceText)
                        Else
                            ' Remplacement simple - le formatage est automatiquement préservé
                            ' car on modifie juste le texte, pas le run ni ses propriétés
                            textElement.Text = textElement.Text.Replace(searchText, replaceText)
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            Debug.WriteLine($"Erreur remplacement paragraphe: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Remplace du texte avec un contenu multiligne
    ''' Si le texte contient des doubles sauts de ligne (vbCrLf & vbCrLf), crée des paragraphes séparés
    ''' Sinon utilise des Break pour les simples sauts de ligne
    ''' PRÉSERVE LE FORMATAGE DU RUN ORIGINAL (gras, italique, couleur, police, etc.)
    ''' </summary>
    Private Sub ReplaceWithMultilineText(run As Run, textElement As Text, searchText As String, replaceText As String)
        Try
            Dim originalText As String = textElement.Text
            Dim beforeText As String = ""
            Dim afterText As String = ""

            Dim idx As Integer = originalText.IndexOf(searchText)
            If idx >= 0 Then
                beforeText = originalText.Substring(0, idx)
                afterText = originalText.Substring(idx + searchText.Length)
            End If

            ' Vérifier si on a des doubles sauts de ligne (= nouveaux paragraphes)
            If replaceText.Contains(vbCrLf & vbCrLf) Then
                ' Créer des paragraphes séparés
                ReplaceWithMultipleParagraphs(run, textElement, searchText, replaceText, beforeText, afterText)
            Else
                ' Utiliser des Break pour les simples sauts de ligne
                ' Supprimer le text element original
                textElement.Remove()

                ' Ajouter le texte avant (avec le même formatage - il est dans le même run)
                If Not String.IsNullOrEmpty(beforeText) Then
                    Dim beforeElement As New Text(beforeText)
                    beforeElement.Space = SpaceProcessingModeValues.Preserve
                    run.Append(beforeElement)
                End If

                ' Ajouter le texte de remplacement avec les breaks
                ' Le formatage est préservé car on reste dans le même run
                Dim lines As String() = replaceText.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                For i As Integer = 0 To lines.Length - 1
                    If i > 0 Then
                        run.Append(New Break())
                    End If
                    Dim lineElement As New Text(lines(i))
                    lineElement.Space = SpaceProcessingModeValues.Preserve
                    run.Append(lineElement)
                Next

                ' Ajouter le texte après (avec le même formatage)
                If Not String.IsNullOrEmpty(afterText) Then
                    Dim afterElement As New Text(afterText)
                    afterElement.Space = SpaceProcessingModeValues.Preserve
                    run.Append(afterElement)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"Erreur remplacement multiligne: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Remplace le placeholder par plusieurs paragraphes séparés
    ''' Chaque paragraphe hérite du formatage du paragraphe original
    ''' </summary>
    Private Sub ReplaceWithMultipleParagraphs(run As Run, textElement As Text, searchText As String, replaceText As String, beforeText As String, afterText As String)
        Try
            Dim parentParagraph As Paragraph = run.Ancestors(Of Paragraph)().FirstOrDefault()
            If parentParagraph Is Nothing Then Return
            
            ' Sauvegarder les propriétés du run (formatage)
            Dim runProps As RunProperties = Nothing
            If run.RunProperties IsNot Nothing Then
                runProps = run.RunProperties.CloneNode(True)
            End If
            
            ' Sauvegarder les propriétés du paragraphe
            Dim paraProps As ParagraphProperties = Nothing
            If parentParagraph.ParagraphProperties IsNot Nothing Then
                paraProps = parentParagraph.ParagraphProperties.CloneNode(True)
            End If
            
            ' Séparer par double saut de ligne
            Dim paragraphs As String() = replaceText.Split(New String() {vbCrLf & vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            
            ' Modifier le paragraphe original avec le premier bloc
            textElement.Remove()
            
            If Not String.IsNullOrEmpty(beforeText) Then
                Dim beforeElement As New Text(beforeText)
                beforeElement.Space = SpaceProcessingModeValues.Preserve
                run.Append(beforeElement)
            End If
            
            ' Premier paragraphe dans le run existant
            If paragraphs.Length > 0 Then
                Dim firstLines As String() = paragraphs(0).Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                For i As Integer = 0 To firstLines.Length - 1
                    If i > 0 Then
                        run.Append(New Break())
                    End If
                    Dim lineElement As New Text(firstLines(i))
                    lineElement.Space = SpaceProcessingModeValues.Preserve
                    run.Append(lineElement)
                Next
            End If
            
            ' Ajouter le texte après dans le premier paragraphe (si c'est le seul)
            If paragraphs.Length = 1 AndAlso Not String.IsNullOrEmpty(afterText) Then
                Dim afterElement As New Text(afterText)
                afterElement.Space = SpaceProcessingModeValues.Preserve
                run.Append(afterElement)
            End If
            
            ' Créer les paragraphes suivants
            Dim insertAfter As OpenXmlElement = parentParagraph
            For pIdx As Integer = 1 To paragraphs.Length - 1
                Dim newPara As New Paragraph()
                
                ' Copier les propriétés du paragraphe
                If paraProps IsNot Nothing Then
                    newPara.Append(paraProps.CloneNode(True))
                End If
                
                ' Créer un nouveau run avec le même formatage
                Dim newRun As New Run()
                If runProps IsNot Nothing Then
                    newRun.Append(runProps.CloneNode(True))
                End If
                
                ' Ajouter le texte avec gestion des sauts de ligne simples
                Dim lines As String() = paragraphs(pIdx).Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                For i As Integer = 0 To lines.Length - 1
                    If i > 0 Then
                        newRun.Append(New Break())
                    End If
                    Dim lineElement As New Text(lines(i))
                    lineElement.Space = SpaceProcessingModeValues.Preserve
                    newRun.Append(lineElement)
                Next
                
                ' Ajouter le texte après au dernier paragraphe
                If pIdx = paragraphs.Length - 1 AndAlso Not String.IsNullOrEmpty(afterText) Then
                    Dim afterElement As New Text(afterText)
                    afterElement.Space = SpaceProcessingModeValues.Preserve
                    newRun.Append(afterElement)
                End If
                
                newPara.Append(newRun)
                
                ' Insérer après le paragraphe précédent
                insertAfter.InsertAfterSelf(newPara)
                insertAfter = newPara
            Next
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur création paragraphes multiples: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Insère un tableau à la place d'un placeholder {xxx}
    ''' </summary>
    Private Sub InsertTableAtPlaceholder(body As Body, placeholder As String, table As Table)
        Try
            For Each paragraph In body.Descendants(Of Paragraph)().ToList()
                Dim paragraphText As String = paragraph.InnerText
                If paragraphText.Contains(placeholder) Then
                    ' Récupérer la police, taille et couleur du paragraphe source
                    Dim fontName As String = Nothing
                    Dim fontSize As String = Nothing
                    Dim fontColor As String = Nothing
                    
                    Dim firstRun = paragraph.Descendants(Of Run)().FirstOrDefault()
                    If firstRun IsNot Nothing AndAlso firstRun.RunProperties IsNot Nothing Then
                        ' Police
                        Dim runFonts = firstRun.RunProperties.GetFirstChild(Of RunFonts)()
                        If runFonts IsNot Nothing Then
                            fontName = runFonts.Ascii
                        End If
                        ' Taille
                        Dim fontSizeEl = firstRun.RunProperties.GetFirstChild(Of FontSize)()
                        If fontSizeEl IsNot Nothing Then
                            fontSize = fontSizeEl.Val
                        End If
                        ' Couleur
                        Dim colorEl = firstRun.RunProperties.GetFirstChild(Of Color)()
                        If colorEl IsNot Nothing Then
                            fontColor = colorEl.Val
                        End If
                    End If
                    
                    ' Appliquer la police, taille et couleur à TOUS les Runs du tableau
                    If Not String.IsNullOrEmpty(fontName) OrElse Not String.IsNullOrEmpty(fontSize) OrElse Not String.IsNullOrEmpty(fontColor) Then
                        For Each tableRun In table.Descendants(Of Run)()
                            ' Créer RunProperties si n'existe pas
                            If tableRun.RunProperties Is Nothing Then
                                tableRun.PrependChild(New RunProperties())
                            End If
                            
                            Dim runProps As RunProperties = tableRun.RunProperties
                            
                            ' Police (si pas déjà définie)
                            If Not String.IsNullOrEmpty(fontName) Then
                                Dim existingFonts = runProps.GetFirstChild(Of RunFonts)()
                                If existingFonts Is Nothing Then
                                    runProps.PrependChild(New RunFonts() With {
                                        .Ascii = fontName,
                                        .HighAnsi = fontName,
                                        .ComplexScript = fontName
                                    })
                                End If
                            End If
                            
                            ' Taille (si pas déjà définie)
                            If Not String.IsNullOrEmpty(fontSize) Then
                                Dim existingSize = runProps.GetFirstChild(Of FontSize)()
                                If existingSize Is Nothing Then
                                    runProps.Append(New FontSize() With {.Val = fontSize})
                                    runProps.Append(New FontSizeComplexScript() With {.Val = fontSize})
                                End If
                            End If
                            
                            ' Couleur (si pas déjà définie)
                            If Not String.IsNullOrEmpty(fontColor) Then
                                Dim existingColor = runProps.GetFirstChild(Of Color)()
                                If existingColor Is Nothing Then
                                    runProps.Append(New Color() With {.Val = fontColor})
                                End If
                            End If
                        Next
                    End If
                    
                    ' Insérer le tableau avant le paragraphe
                    paragraph.InsertBeforeSelf(table)
                    
                    ' Supprimer le paragraphe contenant le placeholder
                    paragraph.Remove()
                    
                    Return ' Un seul remplacement
                End If
            Next
        Catch ex As Exception
            _log.Add($"Erreur insertion tableau: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Insère des segments formatés à la place d'un placeholder {xxx}
    ''' Préserve le formatage Gras/Italique/Souligné du template_paragrahs.docx
    ''' </summary>
    Private Sub InsertFormattedSegmentsAtPlaceholder(body As Body, placeholder As String, segments As List(Of FormattedSegment))
        Try
            For Each paragraph In body.Descendants(Of Paragraph)().ToList()
                Dim paragraphText As String = paragraph.InnerText
                If paragraphText.Contains(placeholder) Then
                    ' Créer un nouveau paragraphe avec les segments formatés
                    Dim newParagraph As New Paragraph()
                    
                    ' Copier les propriétés du paragraphe original (alignement, etc.)
                    If paragraph.ParagraphProperties IsNot Nothing Then
                        newParagraph.ParagraphProperties = CType(paragraph.ParagraphProperties.CloneNode(True), ParagraphProperties)
                    End If
                    
                    ' Obtenir le formatage de base du run original (police, taille, couleur)
                    Dim baseRunProps As RunProperties = Nothing
                    Dim firstRun = paragraph.Descendants(Of Run)().FirstOrDefault()
                    If firstRun IsNot Nothing AndAlso firstRun.RunProperties IsNot Nothing Then
                        baseRunProps = CType(firstRun.RunProperties.CloneNode(True), RunProperties)
                        ' Supprimer le gras du base (sera ajouté selon les segments)
                        If baseRunProps.Bold IsNot Nothing Then baseRunProps.Bold.Remove()
                    End If
                    
                    ' Ajouter chaque segment avec son formatage
                    For Each seg In segments
                        Dim run As New Run()
                        
                        ' Créer les propriétés du run
                        Dim runProps As New RunProperties()
                        
                        ' Copier les propriétés de base (police, taille, couleur)
                        If baseRunProps IsNot Nothing Then
                            For Each child In baseRunProps.ChildElements
                                runProps.Append(child.CloneNode(True))
                            Next
                        End If
                        
                        ' Ajouter Gras si nécessaire
                        If seg.IsBold Then
                            runProps.Append(New Bold())
                        End If
                        
                        If runProps.HasChildren Then
                            run.RunProperties = runProps
                        End If
                        
                        ' Gérer les retours à la ligne dans le texte
                        Dim lines As String() = seg.Text.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                        For i As Integer = 0 To lines.Length - 1
                            If i > 0 Then
                                run.Append(New Break())
                            End If
                            Dim textElement As New Text(lines(i))
                            textElement.Space = SpaceProcessingModeValues.Preserve
                            run.Append(textElement)
                        Next
                        
                        newParagraph.Append(run)
                    Next
                    
                    ' Insérer le nouveau paragraphe avant l'ancien
                    paragraph.InsertBeforeSelf(newParagraph)
                    
                    ' Supprimer le paragraphe original
                    paragraph.Remove()
                    
                    Return ' Un seul remplacement
                End If
            Next
        Catch ex As Exception
            _log.Add($"Erreur insertion segments formatés: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Génère toutes les superbalises
    ''' </summary>
    Private Function GenerateSuperbalises(contractType As String) As Dictionary(Of String, Object)
        Dim superbalises As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        Try
            Dim superGen As New SuperbaliseGenerator(_data, _paragraphReader)
            Dim tableGen As New TableGenerator(_data)

            ' {auteurspart} - avec formatage
            _log.Add("  - Génération {auteurspart}...")
            Dim auteursPartFormatted As List(Of FormattedSegment) = superGen.GenerateAuteursPartFormatted()
            If auteursPartFormatted.Count > 0 Then
                superbalises("auteurspart") = auteursPartFormatted
                _log.Add($"    → {auteursPartFormatted.Count} segments générés")
            Else
                _log.Add("    → Aucun contenu")
            End If

            ' {editeurspart} - avec formatage
            _log.Add("  - Génération {editeurspart}...")
            Dim editeursPartFormatted As List(Of FormattedSegment) = superGen.GenerateEditeursPartFormatted()
            If editeursPartFormatted.Count > 0 Then
                superbalises("editeurspart") = editeursPartFormatted
                _log.Add($"    → {editeursPartFormatted.Count} segments générés")
            Else
                _log.Add("    → Aucun contenu")
            End If

            ' {subpart}
            _log.Add("  - Génération {subpart}...")
            Dim subPart As String = superGen.GenerateSubPart()
            superbalises("subpart") = If(subPart, "")
            If Not String.IsNullOrEmpty(subPart) Then
                _log.Add($"    → {subPart.Split(vbCrLf).Length} lignes générées")
            Else
                _log.Add("    → VIDE (vérifier templates START_SUBS, START_SUBS2, START_EACA)")
            End If

            ' {licpart}
            _log.Add("  - Génération {licpart}...")
            Dim licPart As String = superGen.GenerateLicPart()
            superbalises("licpart") = If(licPart, "")
            If Not String.IsNullOrEmpty(licPart) Then
                _log.Add($"    → {licPart.Split(vbCrLf).Length} lignes générées")
            Else
                _log.Add("    → VIDE")
            End If

            ' {tabcreasplit}
            _log.Add("  - Génération {tabcreasplit}...")
            Dim tabCreaSplit As Table = tableGen.GenerateTabCreaSplit()
            If tabCreaSplit IsNot Nothing Then
                superbalises("tabcreasplit") = tabCreaSplit
                _log.Add("    → Tableau de répartition créé")
            End If

            ' {tabcreasplit2}
            _log.Add("  - Génération {tabcreasplit2}...")
            Dim tabCreaSplit2 As Table = tableGen.GenerateTabCreaSplit2()
            If tabCreaSplit2 IsNot Nothing Then
                superbalises("tabcreasplit2") = tabCreaSplit2
                _log.Add("    → Tableau détaillé créé")
            End If

            ' {tabsignature}
            _log.Add("  - Génération {tabsignature}...")
            Dim tabSignature As Table = tableGen.GenerateTabSignature(contractType)
            If tabSignature IsNot Nothing Then
                superbalises("tabsignature") = tabSignature
                _log.Add("    → Tableau de signatures créé")
            End If

            ' =====================================================
            ' BLOCS NON-SACEM (factorisés)
            ' =====================================================
            
            ' {MENTION_NONSACEM} - Coédition + EDITEUR
            _log.Add("  - Génération {MENTION_NONSACEM}...")
            Dim mentionNonSACEM As String = superGen.GenerateMentionNonSACEM()
            superbalises("MENTION_NONSACEM") = If(mentionNonSACEM, "")
            If Not String.IsNullOrEmpty(mentionNonSACEM) Then
                _log.Add("    → Bloc MENTION_NONSACEM généré")
            End If
            
            ' {LIST_NONSACEM} - Non-signataire + lettre
            _log.Add("  - Génération {LIST_NONSACEM}...")
            Dim listNonSACEM As String = superGen.GenerateListNonSACEM()
            superbalises("LIST_NONSACEM") = If(listNonSACEM, "")
            If Not String.IsNullOrEmpty(listNonSACEM) Then
                _log.Add("    → Bloc LIST_NONSACEM généré")
            End If
            
            ' {OGC_NONSACEM} - Droits collectés OGC
            _log.Add("  - Génération {OGC_NONSACEM}...")
            Dim ogcNonSACEM As String = superGen.GenerateOGCNonSACEM()
            superbalises("OGC_NONSACEM") = If(ogcNonSACEM, "")
            If Not String.IsNullOrEmpty(ogcNonSACEM) Then
                _log.Add("    → Bloc OGC_NONSACEM généré")
            End If
            
            ' {BDO_NONSACEM} - Commentaire BDO
            _log.Add("  - Génération {BDO_NONSACEM}...")
            Dim bdoNonSACEM As String = superGen.GenerateBDONonSACEM()
            superbalises("BDO_NONSACEM") = If(bdoNonSACEM, "")
            If Not String.IsNullOrEmpty(bdoNonSACEM) Then
                _log.Add("    → Bloc BDO_NONSACEM généré")
            End If

            ' =====================================================
            ' BLOC DÉPÔT PARTIEL
            ' =====================================================

            ' {MENTION_PARTIEL} - Phrase dépôt partiel
            _log.Add("  - Génération {MENTION_PARTIEL}...")
            Dim mentionPartiel As String = superGen.GenerateMentionPartiel()
            superbalises("MENTION_PARTIEL") = If(mentionPartiel, "")
            If Not String.IsNullOrEmpty(mentionPartiel) Then
                _log.Add("    → Bloc MENTION_PARTIEL généré")
            End If

        Catch ex As Exception
            _log.Add($"Erreur génération superbalises: {ex.Message}")
        End Try

        Return superbalises
    End Function

    ''' <summary>
    ''' Génère tous les contrats (CCDAA, CCEOM, COED)
    ''' </summary>
    Public Shared Function GenerateAllContracts(data As SACEMData,
                                                 paragraphReader As ParagraphTemplateReader,
                                                 templatesDirectory As String,
                                                 outputDirectory As String) As List(Of String)
        Dim allLogs As New List(Of String)
        Dim generator As New ContractGenerator(data, paragraphReader)

        ' Créer le dossier de sortie s'il n'existe pas
        If Not Directory.Exists(outputDirectory) Then
            Directory.CreateDirectory(outputDirectory)
        End If

        ' Nom de base pour les fichiers (basé sur le titre)
        Dim baseName As String = CleanFileName(data.Titre)
        If String.IsNullOrEmpty(baseName) Then baseName = "Document"

        ' Contrats à générer
        Dim contracts As New List(Of Tuple(Of String, String, String)) From {
            Tuple.Create("CCDAA_template.docx", $"CCDAA_{baseName}.docx", "CCDAA"),
            Tuple.Create("CCEOM_template_univ.docx", $"CCEOM_{baseName}.docx", "CCEOM"),
            Tuple.Create("COED_template_univ.docx", $"COED_{baseName}.docx", "COED")
        }

        For Each contract In contracts
            Dim templatePath As String = Path.Combine(templatesDirectory, contract.Item1)
            Dim outputPath As String = Path.Combine(outputDirectory, contract.Item2)

            If File.Exists(templatePath) Then
                generator.Generate(templatePath, outputPath, contract.Item3)
                allLogs.AddRange(generator.GenerationLog)
            Else
                allLogs.Add($"⚠ Template non trouvé: {contract.Item1}")
            End If
        Next

        Return allLogs
    End Function

    ''' <summary>
    ''' Nettoie un nom de fichier
    ''' </summary>
    Private Shared Function CleanFileName(name As String) As String
        If String.IsNullOrEmpty(name) Then Return ""

        Dim invalid As Char() = Path.GetInvalidFileNameChars()
        Dim cleaned As String = name

        For Each c In invalid
            cleaned = cleaned.Replace(c, "_"c)
        Next

        Return cleaned.Trim()
    End Function
End Class
