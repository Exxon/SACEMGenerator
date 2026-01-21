Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Service de lecture et parsing des fichiers JSON SACEM
''' Conservation stricte de la structure sans transformation
''' </summary>
Public Class SACEMJsonReader
    
    ''' <summary>
    ''' Charge un fichier JSON SACEM
    ''' </summary>
    Public Shared Function LoadFromFile(filePath As String) As SACEMData
        Try
            If Not File.Exists(filePath) Then
                Throw New FileNotFoundException($"Fichier JSON introuvable: {filePath}")
            End If

            Dim jsonContent As String = File.ReadAllText(filePath, System.Text.Encoding.UTF8)
            
            If String.IsNullOrWhiteSpace(jsonContent) Then
                Throw New InvalidDataException("Le fichier JSON est vide")
            End If

            ' Parsing du JSON
            Dim jObject As JObject = JObject.Parse(jsonContent)

            ' Création de l'objet SACEMData
            Dim data As New SACEMData With {
                .RawData = jObject,
                .SourceFileName = Path.GetFileName(filePath),
                .LoadedDate = DateTime.Now
            }

            ' Extraction des champs directs
            data.Titre = GetStringValue(jObject, "Titre")
            data.SousTitre = GetStringValue(jObject, "SousTitre")
            data.Interprete = GetStringValue(jObject, "Interprete")
            data.Duree = GetStringValue(jObject, "Duree")
            data.Genre = GetStringValue(jObject, "Genre")
            data.Date = GetStringValue(jObject, "Date")
            data.ISWC = GetStringValue(jObject, "ISWC")
            data.Lieu = GetStringValue(jObject, "Lieu")
            data.Territoire = GetStringValue(jObject, "Territoire")
            data.Arrangement = GetStringValue(jObject, "Arrangement")
            data.Inegalitaire = GetStringValue(jObject, "Inegalitaire")
            data.Commentaire = GetStringValue(jObject, "Commentaire")
            data.Faita = GetStringValue(jObject, "Faita")
            data.Faitle = GetStringValue(jObject, "Faitle")
            data.Declaration = GetStringValue(jObject, "Declaration")
            data.Format = GetStringValue(jObject, "Format")

            ' Parsing des ayants droit
            If jObject("AyantsDroit") IsNot Nothing Then
                Dim ayantsDroitArray As JArray = CType(jObject("AyantsDroit"), JArray)
                For Each ayantDroitToken As JToken In ayantsDroitArray
                    Dim ayantDroit As AyantDroit = ParseAyantDroit(CType(ayantDroitToken, JObject))
                    data.AyantsDroit.Add(ayantDroit)
                Next
            End If

            Return data

        Catch jEx As JsonException
            Throw New InvalidDataException($"Erreur de format JSON: {jEx.Message}", jEx)
        Catch ex As Exception
            Throw New Exception($"Erreur lors du chargement du JSON: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' Parse un ayant droit depuis un JObject
    ''' </summary>
    Private Shared Function ParseAyantDroit(jObj As JObject) As AyantDroit
        Dim ayant As New AyantDroit()

        ' Identite
        If jObj("Identite") IsNot Nothing Then
            Dim identiteObj As JObject = CType(jObj("Identite"), JObject)
            ayant.Identite.Designation = GetStringValue(identiteObj, "Designation")
            ayant.Identite.Type = GetStringValue(identiteObj, "Type")
            ayant.Identite.Pseudonyme = GetStringValue(identiteObj, "Pseudonyme")
            ayant.Identite.Nom = GetStringValue(identiteObj, "Nom")
            ayant.Identite.Prenom = GetStringValue(identiteObj, "Prenom")
            
            ' Naissance (pour personnes physiques)
            ayant.Identite.Nele = GetStringValue(identiteObj, "Nele")
            ayant.Identite.Nea = GetStringValue(identiteObj, "Nea")
            
            ' Pour les personnes physiques, Genre = MR/MME
            ' Pour les personnes morales, Genre contient la forme juridique (SAS, EURL, etc.)
            Dim genreValue As String = GetStringValue(identiteObj, "Genre")
            
            If ayant.Identite.Type = "Physique" Then
                ayant.Identite.Genre = genreValue  ' MR ou MME
                ayant.Identite.FormeJuridique = ""
            Else
                ' Personne morale : Genre contient la forme juridique
                ayant.Identite.Genre = ""
                ayant.Identite.FormeJuridique = genreValue  ' SAS, EURL, SASU, etc.
            End If
            
            ' Lire aussi FormeJuridique si elle existe explicitement
            Dim formeExplicite As String = GetStringValue(identiteObj, "FormeJuridique")
            If Not String.IsNullOrEmpty(formeExplicite) Then
                ayant.Identite.FormeJuridique = formeExplicite
            End If
            
            ayant.Identite.Capital = GetStringValue(identiteObj, "Capital")
            ayant.Identite.RCS = GetStringValue(identiteObj, "RCS")
            ayant.Identite.Siren = GetStringValue(identiteObj, "Siren")
            ayant.Identite.PrenomRepresentant = GetStringValue(identiteObj, "PrenomRepresentant")
            ' Gérer les 2 orthographes possibles : PrénomRepresentant et PrenomRepresentant
            If String.IsNullOrEmpty(ayant.Identite.PrenomRepresentant) Then
                ayant.Identite.PrenomRepresentant = GetStringValue(identiteObj, "PrénomRepresentant")
            End If
            ayant.Identite.NomRepresentant = GetStringValue(identiteObj, "NomRepresentant")
            ayant.Identite.GenreRepresentant = GetStringValue(identiteObj, "GenreRepresentant")
            ayant.Identite.FonctionRepresentant = GetStringValue(identiteObj, "FonctionRepresentant")
        End If

        ' BDO
        If jObj("BDO") IsNot Nothing Then
            Dim bdoObj As JObject = CType(jObj("BDO"), JObject)
            ayant.BDO.Role = GetStringValue(bdoObj, "Role")
            ayant.BDO.COAD_IPI = GetStringValue(bdoObj, "COAD/IPI")
            ayant.BDO.PH = GetStringValue(bdoObj, "PH")
            ayant.BDO.Lettrage = GetStringValue(bdoObj, "Lettrage")
            ayant.BDO.Managelic = GetStringValue(bdoObj, "Managelic")
            ayant.BDO.Managesub = GetStringValue(bdoObj, "Managesub")
        End If

        ' Adresse
        If jObj("Adresse") IsNot Nothing Then
            Dim adresseObj As JObject = CType(jObj("Adresse"), JObject)
            ayant.Adresse.NumVoie = GetStringValue(adresseObj, "NumVoie")
            ayant.Adresse.TypeVoie = GetStringValue(adresseObj, "TypeVoie")
            ayant.Adresse.NomVoie = GetStringValue(adresseObj, "NomVoie")
            ayant.Adresse.CP = GetStringValue(adresseObj, "CP")
            ayant.Adresse.Ville = GetStringValue(adresseObj, "Ville")
            ayant.Adresse.Pays = GetStringValue(adresseObj, "Pays")
        End If

        ' Contact
        If jObj("Contact") IsNot Nothing Then
            Dim contactObj As JObject = CType(jObj("Contact"), JObject)
            ayant.Contact.Mail = GetStringValue(contactObj, "Mail")
            ayant.Contact.Tel = GetStringValue(contactObj, "Tel")
        End If

        Return ayant
    End Function

    ''' <summary>
    ''' Récupère une valeur string depuis un JObject
    ''' </summary>
    Private Shared Function GetStringValue(jObj As JObject, key As String) As String
        Try
            If jObj(key) IsNot Nothing AndAlso jObj(key).Type <> JTokenType.Null Then
                Return jObj(key).ToString()
            End If
        Catch
            ' Ignorer les erreurs
        End Try
        Return String.Empty
    End Function

    ''' <summary>
    ''' Génère un rapport de structure du JSON chargé
    ''' </summary>
    Public Shared Function GenerateStructureReport(data As SACEMData) As String
        Dim report As New System.Text.StringBuilder()
        
        report.AppendLine("=== RAPPORT DE STRUCTURE JSON SACEM ===")
        report.AppendLine($"Fichier: {data.SourceFileName}")
        report.AppendLine($"Chargé le: {data.LoadedDate:dd/MM/yyyy HH:mm:ss}")
        report.AppendLine()
        
        report.AppendLine("INFORMATIONS GÉNÉRALES:")
        report.AppendLine($"  Titre: {data.Titre}")
        report.AppendLine($"  Interprète: {data.Interprete}")
        report.AppendLine($"  Genre: {data.Genre}")
        report.AppendLine($"  Durée: {data.Duree}")
        report.AppendLine($"  Date: {data.Date}")
        report.AppendLine()
        
        report.AppendLine($"AYANTS DROIT: {data.AyantsDroit.Count}")
        
        ' Compter par rôle
        Dim roleCount As New Dictionary(Of String, Integer)
        For Each ayant In data.AyantsDroit
            Dim role As String = ayant.BDO.Role
            If String.IsNullOrEmpty(role) Then role = "Non défini"
            If roleCount.ContainsKey(role) Then
                roleCount(role) += 1
            Else
                roleCount(role) = 1
            End If
        Next
        
        For Each kvp In roleCount.OrderBy(Function(x) x.Key)
            report.AppendLine($"  - Rôle '{kvp.Key}': {kvp.Value} ayant(s) droit")
        Next
        
        report.AppendLine()
        report.AppendLine("DÉTAIL DES AYANTS DROIT:")
        For i As Integer = 0 To data.AyantsDroit.Count - 1
            Dim ayant As AyantDroit = data.AyantsDroit(i)
            report.AppendLine($"  [{i + 1}] {ayant.Identite.Designation}")
            report.AppendLine($"      Type: {ayant.Identite.Type}")
            report.AppendLine($"      Rôle: {ayant.BDO.Role}")
            report.AppendLine($"      Part (PH): {ayant.BDO.PH}%")
            report.AppendLine($"      Lettrage: {ayant.BDO.Lettrage}")
        Next
        
        Return report.ToString()
    End Function

    ''' <summary>
    ''' Valide la structure du JSON SACEM
    ''' </summary>
    Public Shared Function ValidateStructure(data As SACEMData) As (IsValid As Boolean, ErrorMessage As String)
        Try
            If data Is Nothing Then
                Return (False, "Les données SACEM sont nulles")
            End If

            ' Vérification des champs obligatoires
            Dim missingFields As New List(Of String)

            If String.IsNullOrEmpty(data.Titre) Then missingFields.Add("Titre")
            If String.IsNullOrEmpty(data.Genre) Then missingFields.Add("Genre")
            If String.IsNullOrEmpty(data.Duree) Then missingFields.Add("Duree")

            If missingFields.Count > 0 Then
                Return (False, $"Champs obligatoires manquants: {String.Join(", ", missingFields)}")
            End If

            ' Vérification des ayants droit
            If data.AyantsDroit Is Nothing OrElse data.AyantsDroit.Count = 0 Then
                Return (False, "Aucun ayant droit trouvé dans le JSON")
            End If

            ' Vérification de la cohérence des ayants droit
            For Each ayant In data.AyantsDroit
                If String.IsNullOrEmpty(ayant.Identite.Designation) Then
                    Return (False, "Un ayant droit n'a pas de désignation")
                End If
                If String.IsNullOrEmpty(ayant.BDO.Role) Then
                    Return (False, $"L'ayant droit '{ayant.Identite.Designation}' n'a pas de rôle")
                End If
            Next

            Return (True, String.Empty)

        Catch ex As Exception
            Return (False, $"Erreur lors de la validation: {ex.Message}")
        End Try
    End Function
End Class
