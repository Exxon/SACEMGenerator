Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Linq
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

''' <summary>
''' Représente un segment de texte avec son formatage (GRAS uniquement)
''' </summary>
Public Class FormattedSegment
    Public Property Text As String
    Public Property IsBold As Boolean
    
    Public Sub New(text As String, Optional bold As Boolean = False)
        Me.Text = text
        Me.IsBold = bold
    End Sub
End Class

''' <summary>
''' Lecteur de templates de paragraphes
''' Le GRAS est lu depuis le fichier template_paragrahs.docx
''' Si une balise [xxx] est en gras dans le template, la valeur sera en gras
''' </summary>
Public Class ParagraphTemplateReader
    Private _templatePath As String
    Private _templates As Dictionary(Of String, String)
    Private _boldBalises As Dictionary(Of String, HashSet(Of String)) ' templateName -> liste des balises en gras

    Public Sub New(templatePath As String)
        _templatePath = templatePath
        _templates = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        _boldBalises = New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
    End Sub

    ''' <summary>
    ''' Charge et parse les templates de paragraphes
    ''' Détecte quelles balises sont en GRAS
    ''' </summary>
    Public Function LoadTemplates() As Dictionary(Of String, String)
        Try
            If Not File.Exists(_templatePath) Then
                Throw New FileNotFoundException($"Template de paragraphes introuvable: {_templatePath}")
            End If

            _templates.Clear()
            _boldBalises.Clear()

            ' Lire le document
            Dim fullText As New StringBuilder()
            Dim boldRanges As New List(Of Tuple(Of Integer, Integer)) ' (start, end) des zones en gras

            Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(_templatePath, False)
                If wordDoc.MainDocumentPart Is Nothing Then
                    Throw New InvalidDataException("Le document ne contient pas de partie principale")
                End If

                Dim body As Body = wordDoc.MainDocumentPart.Document.Body
                
                For Each paragraph As Paragraph In body.Descendants(Of Paragraph)()
                    For Each run As Run In paragraph.Descendants(Of Run)()
                        Dim text As String = run.InnerText
                        If String.IsNullOrEmpty(text) Then Continue For
                        
                        Dim startPos As Integer = fullText.Length
                        fullText.Append(text)
                        Dim endPos As Integer = fullText.Length
                        
                        ' Vérifier si ce run est en GRAS
                        Dim isBold As Boolean = False
                        If run.RunProperties IsNot Nothing AndAlso run.RunProperties.Bold IsNot Nothing Then
                            isBold = True
                        End If
                        
                        If isBold Then
                            boldRanges.Add(Tuple.Create(startPos, endPos))
                        End If
                    Next
                    fullText.AppendLine()
                Next
            End Using

            Dim content As String = fullText.ToString()
            
            ' Extraire les sections {START_X}...{END_X}
            Dim pattern As String = "\{START_(.*?)\}(.*?)\{END_\1\}"
            Dim matches As MatchCollection = Regex.Matches(content, pattern, RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            
            For Each match As Match In matches
                Dim templateName As String = match.Groups(1).Value.Trim()
                Dim templateContent As String = match.Groups(2).Value.Trim()
                Dim contentStart As Integer = match.Groups(2).Index
                
                _templates(templateName) = templateContent
                
                ' Trouver les balises [xxx] qui sont en GRAS dans cette section
                Dim boldBalisesForTemplate As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                
                Dim balisePattern As String = "\[([^\]]+)\]"
                Dim baliseMatches As MatchCollection = Regex.Matches(templateContent, balisePattern)
                
                For Each baliseMatch As Match In baliseMatches
                    Dim baliseName As String = baliseMatch.Groups(1).Value
                    Dim baliseStartInContent As Integer = baliseMatch.Index
                    Dim baliseEndInContent As Integer = baliseMatch.Index + baliseMatch.Length
                    
                    ' Position absolue dans le document complet
                    Dim absStart As Integer = contentStart + baliseStartInContent
                    Dim absEnd As Integer = contentStart + baliseEndInContent
                    
                    ' Vérifier si cette balise est dans une zone en gras
                    For Each boldRange In boldRanges
                        ' Si au moins une partie de la balise est en gras, on considère qu'elle est en gras
                        If boldRange.Item1 < absEnd AndAlso boldRange.Item2 > absStart Then
                            boldBalisesForTemplate.Add(baliseName)
                            Exit For
                        End If
                    Next
                Next
                
                _boldBalises(templateName) = boldBalisesForTemplate
                
                Debug.WriteLine($"Template '{templateName}': {boldBalisesForTemplate.Count} balises en gras: {String.Join(", ", boldBalisesForTemplate)}")
            Next

            Debug.WriteLine($"Total de {_templates.Count} template(s) chargé(s)")
            Return _templates

        Catch ex As Exception
            Throw New Exception($"Erreur chargement templates: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' Applique un template et retourne des segments formatés
    ''' Le GRAS est appliqué selon ce qui est défini dans le template Word
    ''' </summary>
    Public Function ApplyTemplateFormatted(templateName As String, ayantDroit As AyantDroit) As List(Of FormattedSegment)
        Try
            If Not _templates.ContainsKey(templateName) Then
                Debug.WriteLine($"Template '{templateName}' non trouvé")
                Return New List(Of FormattedSegment)
            End If
            
            Dim templateText As String = _templates(templateName)
            Dim boldSet As HashSet(Of String) = If(_boldBalises.ContainsKey(templateName), _boldBalises(templateName), New HashSet(Of String))
            
            Return BuildFormattedSegments(templateText, ayantDroit, boldSet)
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur ApplyTemplateFormatted '{templateName}': {ex.Message}")
            Return New List(Of FormattedSegment)
        End Try
    End Function

    ''' <summary>
    ''' Construit les segments formatés
    ''' </summary>
    Private Function BuildFormattedSegments(templateText As String, ayantDroit As AyantDroit, boldSet As HashSet(Of String)) As List(Of FormattedSegment)
        Dim segments As New List(Of FormattedSegment)
        
        Dim pattern As String = "\[([^\]]+)\]"
        Dim matches As MatchCollection = Regex.Matches(templateText, pattern)
        
        Dim currentPos As Integer = 0
        
        For Each match As Match In matches
            ' Texte AVANT la balise (non formaté)
            If match.Index > currentPos Then
                Dim textBefore As String = templateText.Substring(currentPos, match.Index - currentPos)
                If Not String.IsNullOrEmpty(textBefore) Then
                    segments.Add(New FormattedSegment(textBefore, False))
                End If
            End If
            
            ' Obtenir la clé et la valeur
            Dim key As String = match.Groups(1).Value
            Dim value As String = GetValueFromAyantDroit(key, ayantDroit)
            
            ' Le GRAS est déterminé par le template
            Dim isBold As Boolean = boldSet.Contains(key)
            
            ' Ajouter le segment (seulement si valeur non vide)
            If Not String.IsNullOrEmpty(value) Then
                segments.Add(New FormattedSegment(value, isBold))
            End If
            
            currentPos = match.Index + match.Length
        Next
        
        ' Texte restant après la dernière balise
        If currentPos < templateText.Length Then
            Dim textAfter As String = templateText.Substring(currentPos)
            If Not String.IsNullOrEmpty(textAfter) Then
                segments.Add(New FormattedSegment(textAfter, False))
            End If
        End If
        
        Return segments
    End Function

    ''' <summary>
    ''' Applique un template (version texte simple)
    ''' </summary>
    Public Function ApplyTemplate(templateName As String, ayantDroit As AyantDroit) As String
        Try
            If Not _templates.ContainsKey(templateName) Then
                Return String.Empty
            End If

            Dim templateText As String = _templates(templateName)
            Return ReplaceBalises(templateText, ayantDroit)

        Catch ex As Exception
            Debug.WriteLine($"Erreur ApplyTemplate '{templateName}': {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Function ReplaceBalises(text As String, ayantDroit As AyantDroit) As String
        Dim pattern As String = "\[([^\]]+)\]"
        Return Regex.Replace(text, pattern, Function(m)
            Return GetValueFromAyantDroit(m.Groups(1).Value, ayantDroit)
        End Function)
    End Function

    Private Function GetValueFromAyantDroit(key As String, ayantDroit As AyantDroit) As String
        Try
            Select Case key.ToLower()
                ' Identite
                Case "designation" : Return If(ayantDroit.Identite.Designation, "")
                Case "type" : Return If(ayantDroit.Identite.Type, "")
                Case "pseudonyme" : Return If(ayantDroit.Identite.Pseudonyme, "")
                Case "nom" : Return If(ayantDroit.Identite.Nom, "").ToUpper()
                Case "prenom" : Return If(ayantDroit.Identite.Prenom, "")
                Case "genre" : Return If(ayantDroit.Identite.Genre, "")
                Case "formejuridique" : Return If(ayantDroit.Identite.FormeJuridique, "")
                Case "capital" : Return If(ayantDroit.Identite.Capital, "")
                Case "rcs" : Return If(ayantDroit.Identite.RCS, "")
                Case "siren" : Return If(ayantDroit.Identite.Siren, "")
                Case "prenomrepresentant" : Return If(ayantDroit.Identite.PrenomRepresentant, "")
                Case "nomrepresentant" : Return If(ayantDroit.Identite.NomRepresentant, "")
                Case "genrerepresentant" : Return If(ayantDroit.Identite.GenreRepresentant, "")
                Case "civiliterepresentant" : Return GetCivilite(ayantDroit.Identite.GenreRepresentant)
                Case "fonctionrepresentant" : Return If(ayantDroit.Identite.FonctionRepresentant, "")
                Case "nele" : Return If(ayantDroit.Identite.Nele, "")
                Case "nea" : Return If(ayantDroit.Identite.Nea, "")

                ' Balises conditionnelles
                Case "civilite" : Return GetCivilite(ayantDroit.Identite.Genre)
                Case "rolegenre" : Return ConvertRole(ayantDroit.BDO.Role, ayantDroit.Identite.Genre)
                Case "ditpseudonyme"
                    If Not String.IsNullOrEmpty(ayantDroit.Identite.Pseudonyme) Then
                        Return " dit " & ayantDroit.Identite.Pseudonyme
                    End If
                    Return ""
                Case "negenre"
                    If Not String.IsNullOrEmpty(ayantDroit.Identite.Nele) Then
                        Return " " & GetNeGenre(ayantDroit.Identite.Genre)
                    End If
                    Return ""
                Case "lenele"
                    If Not String.IsNullOrEmpty(ayantDroit.Identite.Nele) Then
                        Return " le " & ayantDroit.Identite.Nele
                    End If
                    Return ""
                Case "anea"
                    If Not String.IsNullOrEmpty(ayantDroit.Identite.Nea) Then
                        Return " à " & ayantDroit.Identite.Nea
                    End If
                    Return ""

                ' BDO
                Case "role" : Return ConvertRole(ayantDroit.BDO.Role, ayantDroit.Identite.Genre)
                Case "coad/ipi", "coadipi", "coad_ipi" : Return If(ayantDroit.BDO.COAD_IPI, "")
                Case "ph" : Return If(ayantDroit.BDO.PH, "")
                Case "lettrage" : Return If(ayantDroit.BDO.Lettrage, "")

                ' Adresse
                Case "numvoie" : Return If(ayantDroit.Adresse.NumVoie, "")
                Case "typevoie" : Return If(ayantDroit.Adresse.TypeVoie, "")
                Case "nomvoie" : Return If(ayantDroit.Adresse.NomVoie, "")
                Case "cp" : Return If(ayantDroit.Adresse.CP, "")
                Case "ville" : Return If(ayantDroit.Adresse.Ville, "")
                Case "pays" : Return If(ayantDroit.Adresse.Pays, "")
                Case "adressecomplete" : Return ayantDroit.Adresse.GetAdresseComplete()

                ' Contact
                Case "mail" : Return If(ayantDroit.Contact.Mail, "")
                Case "tel" : Return If(ayantDroit.Contact.Tel, "")

                Case Else
                    Debug.WriteLine($"Clé inconnue: {key}")
                    Return ""
            End Select
        Catch ex As Exception
            Debug.WriteLine($"Erreur GetValueFromAyantDroit '{key}': {ex.Message}")
            Return ""
        End Try
    End Function

    Private Function ConvertRole(role As String, genre As String) As String
        If String.IsNullOrEmpty(role) Then Return ""
        If genre = "MME" Then
            Select Case role.ToUpper()
                Case "A" : Return "d'Autrice"
                Case "C" : Return "de Compositrice"
                Case "AR" : Return "d'Arrangeuse"
                Case "AD" : Return "d'Adaptatrice"
                Case "E" : Return "d'Editrice"
                Case "AC" : Return "d'Autrice-Compositrice"
                Case Else : Return role
            End Select
        Else
            Select Case role.ToUpper()
                Case "A" : Return "d'Auteur"
                Case "C" : Return "de Compositeur"
                Case "AR" : Return "d'Arrangeur"
                Case "AD" : Return "d'Adaptateur"
                Case "E" : Return "d'Editeur"
                Case "AC" : Return "d'Auteur-Compositeur"
                Case Else : Return role
            End Select
        End If
    End Function

    Private Function GetCivilite(genre As String) As String
        Select Case If(genre, "").ToUpper()
            Case "MR", "M", "M." : Return "Monsieur"
            Case "MME", "MLLE", "MS" : Return "Madame"
            Case Else : Return ""
        End Select
    End Function

    Private Function GetNeGenre(genre As String) As String
        Select Case If(genre, "").ToUpper()
            Case "MR", "M", "M." : Return "Né"
            Case "MME", "MLLE", "MS" : Return "Née"
            Case Else : Return "Né"
        End Select
    End Function

    Public ReadOnly Property Templates As Dictionary(Of String, String)
        Get
            Return _templates
        End Get
    End Property

    Public Function HasTemplate(templateName As String) As Boolean
        Return _templates.ContainsKey(templateName)
    End Function
    
    ''' <summary>
    ''' Retourne le contenu brut d'un template
    ''' </summary>
    Public Function GetTemplate(templateName As String) As String
        If _templates.ContainsKey(templateName) Then
            Return _templates(templateName)
        End If
        Return ""
    End Function
End Class
