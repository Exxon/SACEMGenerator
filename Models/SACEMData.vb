Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Modèle de données SACEM - Structure exacte du JSON
''' Conservation stricte sans transformation
''' </summary>
Public Class SACEMData
    ''' <summary>
    ''' Objet JSON brut complet
    ''' </summary>
    Public Property RawData As JObject

    ''' <summary>
    ''' Nom du fichier source
    ''' </summary>
    Public Property SourceFileName As String

    ''' <summary>
    ''' Date de chargement
    ''' </summary>
    Public Property LoadedDate As DateTime

    ' ============================================
    ' CHAMPS DIRECTS (niveau racine)
    ' ============================================
    Public Property Titre As String
    Public Property SousTitre As String
    Public Property Interprete As String
    Public Property Duree As String
    Public Property Genre As String
    Public Property [Date] As String
    Public Property ISWC As String
    Public Property Lieu As String
    Public Property Territoire As String
    Public Property Arrangement As String
    Public Property Inegalitaire As String
    Public Property Commentaire As String
    Public Property Faita As String
    Public Property Faitle As String
    Public Property Declaration As String
    Public Property Format As String

    ''' <summary>
    ''' Liste des ayants droit (structure complète)
    ''' </summary>
    Public Property AyantsDroit As List(Of AyantDroit)

    ''' <summary>
    ''' Constructeur
    ''' </summary>
    Public Sub New()
        LoadedDate = DateTime.Now
        AyantsDroit = New List(Of AyantDroit)
    End Sub

    ''' <summary>
    ''' Récupère une valeur à n'importe quel niveau du JSON
    ''' Support de la notation pointée : "AyantsDroit[0].Identite.Designation"
    ''' </summary>
    Public Function GetValue(path As String) As String
        Try
            If RawData Is Nothing Then Return String.Empty

            ' Support des chemins simples
            If RawData(path) IsNot Nothing Then
                Return RawData(path).ToString()
            End If

            ' Support de la notation pointée avec tableaux
            Dim token As JToken = RawData.SelectToken(path)
            If token IsNot Nothing Then
                Return token.ToString()
            End If

        Catch ex As Exception
            Debug.WriteLine($"Erreur GetValue pour '{path}': {ex.Message}")
        End Try
        Return String.Empty
    End Function

    ''' <summary>
    ''' Vérifie si un chemin existe
    ''' </summary>
    Public Function HasPath(path As String) As Boolean
        Try
            If RawData Is Nothing Then Return False
            Dim token As JToken = RawData.SelectToken(path)
            Return token IsNot Nothing
        Catch
            Return False
        End Try
    End Function
End Class

''' <summary>
''' Représente un ayant droit dans le système SACEM
''' </summary>
Public Class AyantDroit
    Public Property Identite As Identite
    Public Property BDO As BDO
    Public Property Adresse As Adresse
    Public Property Contact As Contact

    Public Sub New()
        Identite = New Identite()
        BDO = New BDO()
        Adresse = New Adresse()
        Contact = New Contact()
    End Sub
End Class

''' <summary>
''' Informations d'identité
''' </summary>
Public Class Identite
    Public Property Designation As String
    Public Property Type As String ' "Physique" ou "Moral"
    
    ' Pour les personnes physiques
    Public Property Pseudonyme As String
    Public Property Nom As String
    Public Property Prenom As String
    Public Property Genre As String ' "MR", "MME"
    Public Property Nele As String ' Date de naissance
    Public Property Nea As String ' Lieu de naissance
    
    ' Société de gestion collective
    Public Property SocieteGestion As String ' "SACEM", etc.
    
    ' Pour les personnes morales
    Public Property FormeJuridique As String ' "SAS", "EURL", "SASU", "EI", "ASS"
    Public Property Capital As String
    Public Property RCS As String
    Public Property Siren As String
    Public Property PrenomRepresentant As String
    Public Property NomRepresentant As String
    Public Property GenreRepresentant As String ' "MR", "MME" pour le représentant
    Public Property FonctionRepresentant As String
End Class

''' <summary>
''' Informations BDO (Bulletin de Déclaration d'Œuvre)
''' </summary>
Public Class BDO
    Public Property Role As String ' "A", "C", "E", "AR", "AD"
    Public Property COAD_IPI As String ' Numéro COAD ou IPI
    Public Property PH As String ' Pourcentage
    Public Property Lettrage As String
    Public Property Managelic As String
    Public Property Managesub As String
    Public Property Signataire As Boolean = True  ' TRUE par défaut si absent du JSON
End Class

''' <summary>
''' Adresse
''' </summary>
Public Class Adresse
    Public Property NumVoie As String
    Public Property TypeVoie As String
    Public Property NomVoie As String
    Public Property CP As String
    Public Property Ville As String
    Public Property Pays As String

    ''' <summary>
    ''' Génère l'adresse complète formatée
    ''' </summary>
    Public Function GetAdresseComplete() As String
        Dim parts As New List(Of String)
        
        If Not String.IsNullOrEmpty(NumVoie) Then parts.Add(NumVoie)
        If Not String.IsNullOrEmpty(TypeVoie) Then parts.Add(TypeVoie)
        If Not String.IsNullOrEmpty(NomVoie) Then parts.Add(NomVoie)
        
        Dim ligne1 As String = String.Join(" ", parts)
        Dim ligne2 As String = $"{CP} {Ville}".Trim()
        
        Dim result As New List(Of String)
        If Not String.IsNullOrEmpty(ligne1) Then result.Add(ligne1)
        If Not String.IsNullOrEmpty(ligne2) Then result.Add(ligne2)
        If Not String.IsNullOrEmpty(Pays) Then result.Add(Pays)
        
        Return String.Join(vbCrLf, result)
    End Function
End Class

''' <summary>
''' Informations de contact
''' </summary>
Public Class Contact
    Public Property Mail As String
    Public Property Tel As String
End Class

''' <summary>
''' Informations de date décomposées
''' </summary>
Public Class DateInfo
    Public Property JJ As String
    Public Property MM As String
    Public Property AAAA As String
    Public Property DateComplete As String

    Public Shared Function Parse(dateString As String) As DateInfo
        Dim info As New DateInfo With {.DateComplete = dateString}
        
        Try
            If String.IsNullOrEmpty(dateString) Then Return info
            
            ' Format attendu : JJ/MM/AAAA ou DD/MM/YYYY
            Dim parts As String() = dateString.Split("/"c)
            If parts.Length = 3 Then
                info.JJ = parts(0)
                info.MM = parts(1)
                info.AAAA = parts(2)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Erreur parsing date '{dateString}': {ex.Message}")
        End Try
        
        Return info
    End Function
End Class

''' <summary>
''' Informations de durée décomposées
''' </summary>
Public Class DureeInfo
    Public Property HH As String
    Public Property MM As String
    Public Property SS As String
    Public Property DureeComplete As String

    Public Shared Function Parse(dureeString As String) As DureeInfo
        Dim info As New DureeInfo With {.DureeComplete = dureeString}
        
        Try
            If String.IsNullOrEmpty(dureeString) Then Return info
            
            ' Enlever les guillemets si présents : "00:02:15"
            dureeString = dureeString.Trim(""""c)
            
            ' Format attendu : HH:MM:SS
            Dim parts As String() = dureeString.Split(":"c)
            If parts.Length = 3 Then
                info.HH = parts(0)
                info.MM = parts(1)
                info.SS = parts(2)
            ElseIf parts.Length = 2 Then
                info.HH = "00"
                info.MM = parts(0)
                info.SS = parts(1)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Erreur parsing durée '{dureeString}': {ex.Message}")
        End Try
        
        Return info
    End Function
End Class
