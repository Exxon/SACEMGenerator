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
    ''' Parse un ayant droit depuis un JObject.
    ''' Supporte deux formats :
    '''   - Léger  : Id + BDO à la racine, identité reenrichie depuis XLSX au chargement
    '''   - Verbeux : Id + BDO à la racine + sous-objets Identite / Adresse / Contact
    ''' </summary>
    Private Shared Function ParseAyantDroit(jObj As JObject) As AyantDroit
        Dim ayant As New AyantDroit()

        ' ── BDO (toujours à la racine) ──
        ayant.BDO.Id        = GetStringValue(jObj, "Id")
        ayant.BDO.Role      = GetStringValue(jObj, "Role")
        ayant.BDO.PH        = GetStringValue(jObj, "PH")
        ayant.BDO.DE        = GetStringValue(jObj, "DE")
        ayant.BDO.DR        = GetStringValue(jObj, "DR")
        ayant.BDO.Lettrage  = GetStringValue(jObj, "Lettrage")
        ayant.BDO.Managelic = GetStringValue(jObj, "Managelic")
        ayant.BDO.Managesub = GetStringValue(jObj, "Managesub")

        Dim sigStr As String = GetStringValue(jObj, "Signataire")
        ayant.BDO.Signataire = (sigStr.ToUpper() = "TRUE" OrElse sigStr = "1")

        ' ── Identité ── sous-objet si format verbeux, sinon déduit de l'Id
        Dim identObj As JObject = TryCast(jObj("Identite"), JObject)
        If identObj IsNot Nothing Then
            ' Format verbeux
            ayant.Identite.Type              = GetStringValue(identObj, "Type")
            ayant.Identite.Designation       = GetStringValue(identObj, "Designation")
            ayant.Identite.Pseudonyme        = GetStringValue(identObj, "Pseudonyme")
            ayant.Identite.Nom               = GetStringValue(identObj, "Nom")
            ayant.Identite.Prenom            = GetStringValue(identObj, "Prenom")
            ayant.Identite.Genre             = GetStringValue(identObj, "Genre")
            ayant.Identite.Nele              = GetStringValue(identObj, "Nele")
            ayant.Identite.Nea               = GetStringValue(identObj, "Nea")
            ayant.Identite.SocieteGestion    = GetStringValue(identObj, "SocieteGestion")
            ayant.Identite.FormeJuridique    = GetStringValue(identObj, "FormeJuridique")
            ayant.Identite.Capital           = GetStringValue(identObj, "Capital")
            ayant.Identite.RCS               = GetStringValue(identObj, "RCS")
            ayant.Identite.Siren             = GetStringValue(identObj, "Siren")
            ayant.Identite.GenreRepresentant    = GetStringValue(identObj, "GenreRepresentant")
            ayant.Identite.PrenomRepresentant   = GetStringValue(identObj, "PrenomRepresentant")
            ayant.Identite.NomRepresentant      = GetStringValue(identObj, "NomRepresentant")
            ayant.Identite.FonctionRepresentant = GetStringValue(identObj, "FonctionRepresentant")
        Else
            ' Format léger — Type déduit du préfixe Id, identité reenrichie depuis XLSX au chargement
            Dim idVal As String = ayant.BDO.Id.Trim().ToUpper()
            If idVal.StartsWith("P") Then
                ayant.Identite.Type = "Physique"
            ElseIf idVal.StartsWith("M") Then
                ayant.Identite.Type = "Moral"
            ElseIf ayant.BDO.Role = "E" OrElse ayant.BDO.Role = "AEC" Then
                ayant.Identite.Type = "Moral"
            Else
                ayant.Identite.Type = "Physique"
            End If
        End If

        ' ── Adresse ── sous-objet si format verbeux
        Dim adrObj As JObject = TryCast(jObj("Adresse"), JObject)
        If adrObj IsNot Nothing Then
            ayant.Adresse.NumVoie  = GetStringValue(adrObj, "NumVoie")
            ayant.Adresse.TypeVoie = GetStringValue(adrObj, "TypeVoie")
            ayant.Adresse.NomVoie  = GetStringValue(adrObj, "NomVoie")
            ayant.Adresse.CP       = GetStringValue(adrObj, "CP")
            ayant.Adresse.Ville    = GetStringValue(adrObj, "Ville")
            ayant.Adresse.Pays     = GetStringValue(adrObj, "Pays")
        End If

        ' ── Contact ── sous-objet si format verbeux
        Dim ctcObj As JObject = TryCast(jObj("Contact"), JObject)
        If ctcObj IsNot Nothing Then
            ayant.Contact.Mail = GetStringValue(ctcObj, "Mail")
            ayant.Contact.Tel  = GetStringValue(ctcObj, "Tel")
        End If

        ' ── COAD/IPI ── champ plat optionnel
        ayant.BDO.COAD_IPI = GetStringValue(jObj, "COAD_IPI")

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
            ' Afficher Designation pour moral, Nom Prenom pour physique
            Dim identifiant As String = If(Not String.IsNullOrEmpty(ayant.Identite.Designation), 
                                           ayant.Identite.Designation, 
                                           $"{ayant.Identite.Prenom} {ayant.Identite.Nom}".Trim())
            report.AppendLine($"  [{i + 1}] {identifiant}")
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
                ' Pour les personnes physiques : Nom ou Prenom requis
                ' Pour les personnes morales : Designation requis
                Dim hasIdentity As Boolean = False
                If ayant.Identite.Type = "Physique" Then
                    hasIdentity = Not String.IsNullOrEmpty(ayant.Identite.Nom) OrElse 
                                  Not String.IsNullOrEmpty(ayant.Identite.Prenom)
                Else
                    hasIdentity = Not String.IsNullOrEmpty(ayant.Identite.Designation)
                End If
                
                If Not hasIdentity Then
                    Return (False, "Un ayant droit n'a pas d'identité (Designation pour moral, Nom/Prenom pour physique)")
                End If
                
                If String.IsNullOrEmpty(ayant.BDO.Role) Then
                    Dim identifiant As String = If(Not String.IsNullOrEmpty(ayant.Identite.Designation), 
                                                   ayant.Identite.Designation, 
                                                   $"{ayant.Identite.Nom} {ayant.Identite.Prenom}".Trim())
                    Return (False, $"L'ayant droit '{identifiant}' n'a pas de rôle")
                End If
            Next

            Return (True, String.Empty)

        Catch ex As Exception
            Return (False, $"Erreur lors de la validation: {ex.Message}")
        End Try
    End Function
End Class
