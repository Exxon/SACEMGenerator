Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Wordprocessing
Imports DocumentFormat.OpenXml.Packaging
Imports System.Globalization

''' <summary>
''' Générateur de tableaux dynamiques Word
''' Reproduit tabcreasplit.py, tabcreasplit2.py, tabsignature.py
''' </summary>
Public Class TableGenerator
    Private ReadOnly _data As SACEMData

    Public Sub New(data As SACEMData)
        _data = data
    End Sub

    ''' <summary>
    ''' Génère un tableau de répartition (tabcreasplit)
    ''' 3 colonnes : Droits d'exécution, Droits radio mécaniques, Droits phonographiques
    ''' Colonne 1 : Nom Prénom dit Pseudonyme (GRAS) + saut + RoleGenre (normal) - AUTO
    ''' Colonnes 2-4 : Pourcentages centrés - Largeur minimum garantie
    ''' Options : Ne pas couper sur 2 pages, police héritée de la balise
    ''' </summary>
    Public Function GenerateTabCreaSplit(Optional fontName As String = Nothing, Optional fontSize As String = Nothing) As Table
        ' Créer le tableau
        Dim table As New Table()

        ' Propriétés du tableau
        Dim tableProperties As New TableProperties()
        Dim tableWidth As New TableWidth() With {.Width = "5000", .Type = TableWidthUnitValues.Pct}
        tableProperties.Append(tableWidth)
        
        ' Layout AUTO pour ajuster les colonnes au contenu
        Dim tableLayout As New TableLayout() With {.Type = TableLayoutValues.Autofit}
        tableProperties.Append(tableLayout)
        
        table.Append(tableProperties)
        
        ' Grille du tableau (sans largeur fixe - auto)
        Dim tableGrid As New TableGrid()
        tableGrid.Append(New GridColumn())  ' Colonne 1 - auto
        tableGrid.Append(New GridColumn())  ' Colonne 2
        tableGrid.Append(New GridColumn())  ' Colonne 3
        tableGrid.Append(New GridColumn())  ' Colonne 4
        table.Append(tableGrid)

        ' Ligne d'en-tête (centrée) - NE PAS COUPER
        Dim headerRow As New TableRow()
        headerRow.Append(New TableRowProperties(New CantSplit()))
        headerRow.Append(CreateCellAutoWidth("", True, False, fontName, fontSize))
        headerRow.Append(CreateCellMinWidth("Droits d'exécution publique", True, True, fontName, fontSize, 1500))
        headerRow.Append(CreateCellMinWidth("Droits radio mécaniques", True, True, fontName, fontSize, 1500))
        headerRow.Append(CreateCellMinWidth("Droits phonographiques", True, True, fontName, fontSize, 1500))
        table.Append(headerRow)

        ' Collecter les créateurs avec leurs infos (sans doublons, combinant A+C)
        ' FILTRE NON-SACEM : seuls les membres SACEM sont inclus
        Dim createursInfo As New Dictionary(Of String, CreatorInfo)(StringComparer.OrdinalIgnoreCase)

        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "E" Then Continue For
            If Not IsSACEM(ayant) Then Continue For ' Exclure NON-SACEM
            
            Dim key As String = GetAyantKey(ayant)
            Dim ph As Double
            Double.TryParse(ayant.BDO.PH, NumberStyles.Any, CultureInfo.InvariantCulture, ph)
            
            If createursInfo.ContainsKey(key) Then
                ' Cumuler le PH
                createursInfo(key).PH += ph
                ' Combiner les rôles A+C
                Dim existingRole As String = createursInfo(key).Role
                If (existingRole = "A" AndAlso ayant.BDO.Role = "C") OrElse 
                   (existingRole = "C" AndAlso ayant.BDO.Role = "A") Then
                    createursInfo(key).Role = "AC"
                End If
            Else
                createursInfo(key) = New CreatorInfo With {
                    .Nom = ayant.Identite.Nom,
                    .Prenom = ayant.Identite.Prenom,
                    .Pseudonyme = ayant.Identite.Pseudonyme,
                    .Genre = ayant.Identite.Genre,
                    .Role = ayant.BDO.Role,
                    .PH = ph
                }
            End If
        Next

        ' Calculer la somme totale
        Dim sommeTotal As Double = createursInfo.Values.Sum(Function(c) c.PH)
        If sommeTotal = 0 Then sommeTotal = 1 ' Éviter division par zéro

        ' Lignes de données
        Dim sommeCol1, sommeCol2, sommeCol3 As Double

        For Each kvp In createursInfo
            Dim info As CreatorInfo = kvp.Value
            Dim phNorm As Double = (info.PH * 100) / sommeTotal

            ' Ligne de données - NE PAS COUPER
            Dim dataRow As New TableRow()
            dataRow.Append(New TableRowProperties(New CantSplit()))
            ' Colonne 1 : Nom Prénom dit Pseudo (GRAS) + saut + Rôle (normal)
            dataRow.Append(CreateCreatorCellNoWrap(info, fontName, fontSize))
            dataRow.Append(CreateCellMinWidth($"{phNorm:F2}%", False, True, fontName, fontSize, 1500))
            dataRow.Append(CreateCellMinWidth($"{phNorm:F2}%", False, True, fontName, fontSize, 1500))
            dataRow.Append(CreateCellMinWidth($"{phNorm:F2}%", False, True, fontName, fontSize, 1500))
            table.Append(dataRow)

            sommeCol1 += phNorm
            sommeCol2 += phNorm
            sommeCol3 += phNorm
        Next

        ' Ligne TOTAL (centrée) - NE PAS COUPER
        Dim totalRow As New TableRow()
        totalRow.Append(New TableRowProperties(New CantSplit()))
        totalRow.Append(CreateCellAutoWidth("TOTAL", True, False, fontName, fontSize))
        totalRow.Append(CreateCellMinWidth($"{sommeCol1:F2}%", True, True, fontName, fontSize, 1500))
        totalRow.Append(CreateCellMinWidth($"{sommeCol2:F2}%", True, True, fontName, fontSize, 1500))
        totalRow.Append(CreateCellMinWidth($"{sommeCol3:F2}%", True, True, fontName, fontSize, 1500))
        table.Append(totalRow)

        Return table
    End Function

    ''' <summary>
    ''' Génère un tableau de répartition détaillé (tabcreasplit2)
    ''' 7 colonnes : Art. 12-1-1° à 12-1-6°
    ''' Colonne 1 : Nom Prénom dit Pseudonyme (GRAS) + saut + RoleGenre (normal) - AUTO
    ''' Colonnes 2-7 : Pourcentages centrés - Largeur minimum garantie
    ''' Options : Ne pas couper sur 2 pages, police héritée de la balise
    ''' </summary>
    Public Function GenerateTabCreaSplit2(Optional fontName As String = Nothing, Optional fontSize As String = Nothing) As Table
        Dim table As New Table()

        ' Propriétés du tableau
        Dim tableProperties As New TableProperties()
        Dim tableWidth As New TableWidth() With {.Width = "5000", .Type = TableWidthUnitValues.Pct}
        tableProperties.Append(tableWidth)
        
        ' Layout AUTO pour ajuster les colonnes au contenu
        Dim tableLayout As New TableLayout() With {.Type = TableLayoutValues.Autofit}
        tableProperties.Append(tableLayout)
        
        table.Append(tableProperties)
        
        ' Grille du tableau (sans largeur fixe - auto)
        Dim tableGrid As New TableGrid()
        tableGrid.Append(New GridColumn())  ' Colonne 1 - auto
        For i As Integer = 1 To 6
            tableGrid.Append(New GridColumn())
        Next
        table.Append(tableGrid)

        ' Ligne d'en-tête (centrée) - NE PAS COUPER
        Dim headerRow As New TableRow()
        headerRow.Append(New TableRowProperties(New CantSplit()))
        headerRow.Append(CreateCellAutoWidth("", True, False, fontName, fontSize))
        For i As Integer = 1 To 6
            headerRow.Append(CreateCellMinWidth($"Art. 12-1-{i}°", True, True, fontName, fontSize, 900))
        Next
        table.Append(headerRow)

        ' Collecter les créateurs avec leurs infos (sans doublons, combinant A+C)
        ' FILTRE NON-SACEM : seuls les membres SACEM sont inclus
        Dim createursInfo As New Dictionary(Of String, CreatorInfo)(StringComparer.OrdinalIgnoreCase)

        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role = "E" Then Continue For
            If Not IsSACEM(ayant) Then Continue For ' Exclure NON-SACEM
            
            Dim key As String = GetAyantKey(ayant)
            Dim ph As Double
            Double.TryParse(ayant.BDO.PH, NumberStyles.Any, CultureInfo.InvariantCulture, ph)
            
            If createursInfo.ContainsKey(key) Then
                createursInfo(key).PH += ph
                Dim existingRole As String = createursInfo(key).Role
                If (existingRole = "A" AndAlso ayant.BDO.Role = "C") OrElse 
                   (existingRole = "C" AndAlso ayant.BDO.Role = "A") Then
                    createursInfo(key).Role = "AC"
                End If
            Else
                createursInfo(key) = New CreatorInfo With {
                    .Nom = ayant.Identite.Nom,
                    .Prenom = ayant.Identite.Prenom,
                    .Pseudonyme = ayant.Identite.Pseudonyme,
                    .Genre = ayant.Identite.Genre,
                    .Role = ayant.BDO.Role,
                    .PH = ph
                }
            End If
        Next

        ' Calculer la somme totale
        Dim sommeTotal As Double = createursInfo.Values.Sum(Function(c) c.PH)
        If sommeTotal = 0 Then sommeTotal = 1

        ' Lignes de données
        Dim sommes(5) As Double

        For Each kvp In createursInfo
            Dim info As CreatorInfo = kvp.Value
            Dim phNorm As Double = (info.PH * 100) / sommeTotal

            ' Ligne de données - NE PAS COUPER
            Dim dataRow As New TableRow()
            dataRow.Append(New TableRowProperties(New CantSplit()))
            ' Colonne 1 : Nom Prénom dit Pseudo (GRAS) + saut + Rôle (normal)
            dataRow.Append(CreateCreatorCellNoWrap(info, fontName, fontSize))
            
            For i As Integer = 0 To 5
                dataRow.Append(CreateCellMinWidth($"{phNorm:F2}%", False, True, fontName, fontSize, 900))
                sommes(i) += phNorm
            Next
            
            table.Append(dataRow)
        Next

        ' Ligne TOTAL (centrée) - NE PAS COUPER
        Dim totalRow As New TableRow()
        totalRow.Append(New TableRowProperties(New CantSplit()))
        totalRow.Append(CreateCellAutoWidth("TOTAL", True, False, fontName, fontSize))
        For i As Integer = 0 To 5
            totalRow.Append(CreateCellMinWidth($"{sommes(i):F2}%", True, True, fontName, fontSize, 900))
        Next
        table.Append(totalRow)

        Return table
    End Function

    ''' <summary>
    ''' Génère un tableau de signatures
    ''' 4 colonnes, ordre : auteurs/compositeurs PUIS éditeurs (cohérent avec auteurspart et editeurspart)
    ''' Format : Prénom NOM ou DESIGNATION + saut + Rôle + 5 sauts de ligne
    ''' Pas de doublons
    ''' </summary>
    Public Function GenerateTabSignature(typeTemplate As String) As Table
        Dim signatures As New List(Of SignatureInfo)
        
        ' =============================================
        ' 1. D'ABORD LES AUTEURS/COMPOSITEURS (ordre de {auteurspart})
        ' FILTRE NON-SACEM : seuls les membres SACEM signent
        ' =============================================
        If typeTemplate <> "COED" Then
            Dim auteursDict As New Dictionary(Of String, SignatureInfo)(StringComparer.OrdinalIgnoreCase)
            Dim rolesParAuteur As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
            
            For Each ayant In _data.AyantsDroit
                Dim role As String = ayant.BDO.Role
                If role = "E" Then Continue For ' Ignorer les éditeurs pour cette partie
                If Not IsSACEM(ayant) Then Continue For ' Exclure NON-SACEM
                
                ' Clé : Prénom NOM
                Dim prenom As String = If(ayant.Identite.Prenom, "").Trim()
                Dim nom As String = If(ayant.Identite.Nom, "").ToUpper().Trim()
                Dim displayName As String = $"{prenom} {nom}".Trim()
                
                If String.IsNullOrEmpty(displayName) Then
                    displayName = If(ayant.Identite.Designation, "").Trim()
                End If
                
                Dim key As String = NormalizeKey(displayName)
                If String.IsNullOrEmpty(key) Then Continue For
                
                If Not auteursDict.ContainsKey(key) Then
                    rolesParAuteur(key) = New List(Of String)
                    auteursDict(key) = New SignatureInfo With {
                        .Designation = displayName,
                        .Genre = ayant.Identite.Genre
                    }
                End If
                
                If Not rolesParAuteur(key).Contains(role) Then
                    rolesParAuteur(key).Add(role)
                End If
            Next
            
            ' Déterminer le rôle combiné et ajouter à la liste
            For Each kvp In auteursDict
                Dim roles As List(Of String) = rolesParAuteur(kvp.Key)
                Dim genre As String = kvp.Value.Genre
                
                If roles.Contains("A") AndAlso roles.Contains("C") Then
                    kvp.Value.Role = ConvertRoleForSignature("AC", genre)
                ElseIf roles.Contains("A") Then
                    kvp.Value.Role = ConvertRoleForSignature("A", genre)
                ElseIf roles.Contains("C") Then
                    kvp.Value.Role = ConvertRoleForSignature("C", genre)
                ElseIf roles.Contains("AR") Then
                    kvp.Value.Role = ConvertRoleForSignature("AR", genre)
                ElseIf roles.Contains("AD") Then
                    kvp.Value.Role = ConvertRoleForSignature("AD", genre)
                Else
                    kvp.Value.Role = ConvertRoleForSignature(roles.FirstOrDefault(), genre)
                End If
                
                signatures.Add(kvp.Value)
            Next
        End If
        
        ' =============================================
        ' 2. ENSUITE LES ÉDITEURS (ordre de {editeurspart})
        ' FILTRE NON-SACEM : seuls les membres SACEM signent
        ' =============================================
        Dim editeursDict As New Dictionary(Of String, SignatureInfo)(StringComparer.OrdinalIgnoreCase)
        
        For Each ayant In _data.AyantsDroit
            If ayant.BDO.Role <> "E" Then Continue For
            If Not IsSACEM(ayant) Then Continue For ' Exclure NON-SACEM
            
            Dim displayName As String = If(ayant.Identite.Designation, "").Trim()
            Dim key As String = NormalizeKey(displayName)
            
            If String.IsNullOrEmpty(key) Then Continue For
            
            If Not editeursDict.ContainsKey(key) Then
                editeursDict(key) = New SignatureInfo With {
                    .Designation = displayName,
                    .Role = ConvertRoleForSignature("E", ayant.Identite.Genre),
                    .Genre = ayant.Identite.Genre
                }
            End If
        Next
        
        ' Ajouter les éditeurs à la liste
        For Each kvp In editeursDict
            signatures.Add(kvp.Value)
        Next
        
        ' =============================================
        ' 3. CRÉER LE TABLEAU (4 colonnes)
        ' =============================================
        Dim table As New Table()
        
        ' Propriétés du tableau (SANS BORDURES)
        Dim tableProperties As New TableProperties()
        Dim tableWidth As New TableWidth() With {.Width = "5000", .Type = TableWidthUnitValues.Pct}
        tableProperties.Append(tableWidth)
        
        ' Bordures du tableau = AUCUNE
        Dim tableBorders As New TableBorders()
        tableBorders.Append(New TopBorder() With {.Val = BorderValues.Nil, .Size = 0})
        tableBorders.Append(New BottomBorder() With {.Val = BorderValues.Nil, .Size = 0})
        tableBorders.Append(New LeftBorder() With {.Val = BorderValues.Nil, .Size = 0})
        tableBorders.Append(New RightBorder() With {.Val = BorderValues.Nil, .Size = 0})
        tableBorders.Append(New InsideHorizontalBorder() With {.Val = BorderValues.Nil, .Size = 0})
        tableBorders.Append(New InsideVerticalBorder() With {.Val = BorderValues.Nil, .Size = 0})
        tableProperties.Append(tableBorders)
        table.Append(tableProperties)
        
        ' Grille du tableau (obligatoire)
        Dim tableGrid As New TableGrid()
        For i As Integer = 0 To 3
            tableGrid.Append(New GridColumn())
        Next
        table.Append(tableGrid)
        
        ' Remplir le tableau (4 colonnes par ligne)
        Dim nbRows As Integer = CInt(System.Math.Ceiling(signatures.Count / 4.0))
        Dim sigIndex As Integer = 0
        
        For r As Integer = 0 To nbRows - 1
            Dim row As New TableRow()
            
            For c As Integer = 0 To 3
                If sigIndex < signatures.Count Then
                    Dim sig As SignatureInfo = signatures(sigIndex)
                    ' Format : Prénom NOM + saut + Rôle + 5 sauts de ligne
                    Dim cellContent As String = $"{sig.Designation}{vbCrLf}{sig.Role}{vbCrLf}{vbCrLf}{vbCrLf}{vbCrLf}{vbCrLf}"
                    row.Append(CreateSignatureCellNoBorder(cellContent))
                Else
                    row.Append(CreateSignatureCellNoBorder(""))
                End If
                sigIndex += 1
            Next
            
            table.Append(row)
        Next
        
        Return table
    End Function
    
    ''' <summary>
    ''' Normalise une clé pour la comparaison (majuscules, sans espaces multiples, sans accents)
    ''' </summary>
    Private Function NormalizeKey(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        
        ' Majuscules
        Dim result As String = text.ToUpper().Trim()
        
        ' Remplacer les espaces multiples par un seul
        While result.Contains("  ")
            result = result.Replace("  ", " ")
        End While
        
        Return result
    End Function

    ''' <summary>
    ''' Obtient une clé unique pour un ayant droit
    ''' Pour les personnes morales : Designation
    ''' Pour les personnes physiques : Nom + Prenom
    ''' </summary>
    Private Function GetAyantKey(ayant As AyantDroit) As String
        If Not String.IsNullOrEmpty(ayant.Identite.Designation) Then
            Return NormalizeKey(ayant.Identite.Designation)
        Else
            ' Pour les personnes physiques, utiliser Nom + Prenom
            Return NormalizeKey($"{ayant.Identite.Nom} {ayant.Identite.Prenom}")
        End If
    End Function
    
    ''' <summary>
    ''' Crée une cellule pour le tableau de signatures (SANS AUCUNE BORDURE, CONTENU CENTRÉ)
    ''' </summary>
    Private Function CreateSignatureCellNoBorder(content As String) As TableCell
        Dim cell As New TableCell()
        
        ' Propriétés de cellule (SANS BORDURES)
        Dim cellProperties As New TableCellProperties()
        
        ' Bordures de cellule = AUCUNE (Nil)
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Nil, .Size = 0})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Nil, .Size = 0})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Nil, .Size = 0})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Nil, .Size = 0})
        cellProperties.Append(cellBorders)
        
        cell.Append(cellProperties)
        
        ' Créer un paragraphe avec le contenu multiligne ET ALIGNEMENT CENTRÉ
        Dim paragraph As New Paragraph()
        
        ' Alignement centré
        Dim paragraphProperties As New ParagraphProperties()
        paragraphProperties.Append(New Justification() With {.Val = JustificationValues.Center})
        paragraph.Append(paragraphProperties)
        
        Dim run As New Run()
        
        ' Gérer les sauts de ligne
        Dim lines As String() = content.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
        For i As Integer = 0 To lines.Length - 1
            If i > 0 Then
                run.Append(New Break())
            End If
            Dim textElement As New Text(lines(i))
            textElement.Space = SpaceProcessingModeValues.Preserve
            run.Append(textElement)
        Next
        
        paragraph.Append(run)
        cell.Append(paragraph)
        
        Return cell
    End Function
    
    ''' <summary>
    ''' Construit le contenu d'une colonne de signatures
    ''' Format : NOM Prénom (saut) Rôle (2 sauts entre personnes)
    ''' </summary>
    Private Function BuildSignatureContent(signatures As List(Of SignatureInfo)) As String
        Dim sb As New System.Text.StringBuilder()
        
        For i As Integer = 0 To signatures.Count - 1
            Dim sig As SignatureInfo = signatures(i)
            
            ' NOM Prénom
            sb.Append(sig.Designation)
            sb.Append(vbCrLf)
            
            ' Rôle
            sb.Append(sig.Role)
            
            ' 2 sauts de ligne entre les personnes (sauf la dernière)
            If i < signatures.Count - 1 Then
                sb.Append(vbCrLf)
                sb.Append(vbCrLf)
            End If
        Next
        
        Return sb.ToString()
    End Function
    
    ''' <summary>
    ''' Crée une cellule pour le tableau de signatures (sans bordures, avec sauts de ligne)
    ''' </summary>
    Private Function CreateSignatureCell(content As String) As TableCell
        Dim cell As New TableCell()
        
        ' Propriétés de cellule (SANS BORDURES)
        Dim cellProperties As New TableCellProperties()
        Dim cellWidth As New TableCellWidth() With {.Width = "2500", .Type = TableWidthUnitValues.Pct}
        cellProperties.Append(cellWidth)
        
        ' Bordures de cellule = None
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Nil})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Nil})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Nil})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Nil})
        cellProperties.Append(cellBorders)
        
        cell.Append(cellProperties)
        
        ' Créer un paragraphe avec le contenu multiligne
        Dim paragraph As New Paragraph()
        Dim run As New Run()
        
        ' Gérer les sauts de ligne
        Dim lines As String() = content.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
        For i As Integer = 0 To lines.Length - 1
            If i > 0 Then
                run.Append(New Break())
            End If
            Dim textElement As New Text(lines(i))
            textElement.Space = SpaceProcessingModeValues.Preserve
            run.Append(textElement)
        Next
        
        paragraph.Append(run)
        cell.Append(paragraph)
        
        Return cell
    End Function

    ''' <summary>
    ''' Crée une cellule de tableau
    ''' </summary>
    Private Function CreateCell(text As String, Optional isBold As Boolean = False, Optional isCenter As Boolean = False) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' Alignement
        If isCenter Then
            Dim paragraphProperties As New ParagraphProperties()
            paragraphProperties.Append(New Justification() With {.Val = JustificationValues.Center})
            paragraph.Append(paragraphProperties)
        End If

        ' Run
        Dim run As New Run()
        If isBold Then
            Dim runProperties As New RunProperties()
            runProperties.Append(New Bold())
            run.Append(runProperties)
        End If
        run.Append(New Text(text))
        
        paragraph.Append(run)
        cell.Append(paragraph)

        Return cell
    End Function

    ''' <summary>
    ''' Convertit un rôle pour la signature
    ''' </summary>
    Private Function ConvertRoleForSignature(role As String, genre As String) As String
        If String.IsNullOrEmpty(role) Then Return ""
        
        If genre = "MME" Then
            Select Case role.ToUpper()
                Case "A" : Return "Autrice"
                Case "C" : Return "Compositrice"
                Case "AR" : Return "Arrangeuse"
                Case "AD" : Return "Adaptatrice"
                Case "AC" : Return "Autrice et Compositrice"
                Case "E" : Return "Editeur"
                Case Else : Return role
            End Select
        Else
            Select Case role.ToUpper()
                Case "A" : Return "Auteur"
                Case "C" : Return "Compositeur"
                Case "AR" : Return "Arrangeur"
                Case "AD" : Return "Adaptateur"
                Case "AC" : Return "Auteur et Compositeur"
                Case "E" : Return "Editeur"
                Case Else : Return role
            End Select
        End If
    End Function

    ''' <summary>
    ''' Crée une cellule avec largeur automatique (s'adapte au contenu)
    ''' </summary>
    Private Function CreateCellAutoWidth(text As String, isBold As Boolean, isCenter As Boolean, fontName As String, fontSize As String) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        
        ' Largeur auto
        cellProperties.Append(New TableCellWidth() With {.Type = TableWidthUnitValues.Auto})
        
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' Alignement
        If isCenter Then
            Dim paragraphProperties As New ParagraphProperties()
            paragraphProperties.Append(New Justification() With {.Val = JustificationValues.Center})
            paragraph.Append(paragraphProperties)
        End If

        ' Run
        Dim run As New Run()
        Dim runProperties As New RunProperties()
        
        If isBold Then
            runProperties.Append(New Bold())
        End If
        
        If Not String.IsNullOrEmpty(fontName) Then
            runProperties.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
        End If
        If Not String.IsNullOrEmpty(fontSize) Then
            runProperties.Append(New FontSize() With {.Val = fontSize})
        End If
        
        If runProperties.HasChildren Then
            run.Append(runProperties)
        End If
        
        run.Append(New Text(text))
        
        paragraph.Append(run)
        cell.Append(paragraph)

        Return cell
    End Function
    
    ''' <summary>
    ''' Crée une cellule avec largeur minimum garantie (en twips, 1440 twips = 1 pouce)
    ''' </summary>
    Private Function CreateCellMinWidth(text As String, isBold As Boolean, isCenter As Boolean, fontName As String, fontSize As String, minWidthTwips As Integer) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        
        ' Largeur minimum en DXA (twips)
        cellProperties.Append(New TableCellWidth() With {
            .Width = minWidthTwips.ToString(),
            .Type = TableWidthUnitValues.Dxa
        })
        
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' Alignement
        If isCenter Then
            Dim paragraphProperties As New ParagraphProperties()
            paragraphProperties.Append(New Justification() With {.Val = JustificationValues.Center})
            paragraph.Append(paragraphProperties)
        End If

        ' Run
        Dim run As New Run()
        Dim runProperties As New RunProperties()
        
        If isBold Then
            runProperties.Append(New Bold())
        End If
        
        If Not String.IsNullOrEmpty(fontName) Then
            runProperties.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
        End If
        If Not String.IsNullOrEmpty(fontSize) Then
            runProperties.Append(New FontSize() With {.Val = fontSize})
        End If
        
        If runProperties.HasChildren Then
            run.Append(runProperties)
        End If
        
        run.Append(New Text(text))
        
        paragraph.Append(run)
        cell.Append(paragraph)

        Return cell
    End Function

    ''' <summary>
    ''' Clone un ayant droit
    ''' </summary>
    Private Function CloneAyantDroit(source As AyantDroit) As AyantDroit
        Dim clone As New AyantDroit()
        ' Copie simple - dans production, utiliser une vraie deep copy
        clone.Identite.Designation = source.Identite.Designation
        clone.Identite.Nom = source.Identite.Nom
        clone.Identite.Prenom = source.Identite.Prenom
        clone.Identite.Genre = source.Identite.Genre
        clone.BDO.Role = source.BDO.Role
        clone.BDO.PH = source.BDO.PH
        Return clone
    End Function

    ''' <summary>
    ''' Formate le nom du créateur : Nom Prénom dit Pseudonyme + saut + RoleGenre
    ''' </summary>
    Private Function FormatCreatorName(info As CreatorInfo) As String
        Dim sb As New System.Text.StringBuilder()
        
        ' Nom Prénom
        Dim nom As String = If(info.Nom, "").ToUpper()
        Dim prenom As String = If(info.Prenom, "")
        sb.Append($"{nom} {prenom}".Trim())
        
        ' dit Pseudonyme (si non vide)
        If Not String.IsNullOrEmpty(info.Pseudonyme) Then
            sb.Append($" dit {info.Pseudonyme}")
        End If
        
        ' Saut de ligne + RoleGenre
        sb.Append(vbCrLf)
        sb.Append(ConvertRoleGenre(info.Role, info.Genre))
        
        Return sb.ToString()
    End Function
    
    ''' <summary>
    ''' Crée une cellule pour le créateur : Nom Prénom dit Pseudo (GRAS) + saut + Rôle (normal)
    ''' Avec NoWrap pour que le nom tienne sur une ligne
    ''' </summary>
    Private Function CreateCreatorCellNoWrap(info As CreatorInfo, Optional fontName As String = Nothing, Optional fontSize As String = Nothing) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule avec bordures ET NoWrap
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        
        ' NoWrap = pas de retour à la ligne automatique
        cellProperties.Append(New NoWrap())
        
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' === LIGNE 1 : Nom Prénom dit Pseudo - EN GRAS ===
        Dim runBold As New Run()
        Dim runPropsBold As New RunProperties()
        runPropsBold.Append(New Bold())
        ' Appliquer la police si spécifiée
        If Not String.IsNullOrEmpty(fontName) Then
            runPropsBold.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
        End If
        If Not String.IsNullOrEmpty(fontSize) Then
            runPropsBold.Append(New FontSize() With {.Val = fontSize})
        End If
        runBold.Append(runPropsBold)
        
        ' Construire le texte : Nom Prénom dit Pseudonyme
        Dim nom As String = If(info.Nom, "").ToUpper()
        Dim prenom As String = If(info.Prenom, "")
        Dim ligne1 As String = $"{nom} {prenom}".Trim()
        
        If Not String.IsNullOrEmpty(info.Pseudonyme) Then
            ligne1 &= $" dit {info.Pseudonyme}"
        End If
        
        Dim text1 As New Text(ligne1)
        text1.Space = SpaceProcessingModeValues.Preserve
        runBold.Append(text1)
        paragraph.Append(runBold)
        
        ' === SAUT DE LIGNE ===
        Dim runBreak As New Run()
        runBreak.Append(New Break())
        paragraph.Append(runBreak)
        
        ' === LIGNE 2 : Rôle - NORMAL ===
        Dim runNormal As New Run()
        ' Appliquer la police si spécifiée
        If Not String.IsNullOrEmpty(fontName) OrElse Not String.IsNullOrEmpty(fontSize) Then
            Dim runPropsNormal As New RunProperties()
            If Not String.IsNullOrEmpty(fontName) Then
                runPropsNormal.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
            End If
            If Not String.IsNullOrEmpty(fontSize) Then
                runPropsNormal.Append(New FontSize() With {.Val = fontSize})
            End If
            runNormal.Append(runPropsNormal)
        End If
        
        Dim ligne2 As String = ConvertRoleGenre(info.Role, info.Genre)
        Dim text2 As New Text(ligne2)
        text2.Space = SpaceProcessingModeValues.Preserve
        runNormal.Append(text2)
        paragraph.Append(runNormal)
        
        cell.Append(paragraph)
        Return cell
    End Function

    ''' <summary>
    ''' Crée une cellule pour le créateur : Nom Prénom dit Pseudo (GRAS) + saut + Rôle (normal)
    ''' </summary>
    Private Function CreateCreatorCell(info As CreatorInfo, Optional fontName As String = Nothing, Optional fontSize As String = Nothing) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule avec bordures
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' === LIGNE 1 : Nom Prénom dit Pseudo - EN GRAS ===
        Dim runBold As New Run()
        Dim runPropsBold As New RunProperties()
        runPropsBold.Append(New Bold())
        ' Appliquer la police si spécifiée
        If Not String.IsNullOrEmpty(fontName) Then
            runPropsBold.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
        End If
        If Not String.IsNullOrEmpty(fontSize) Then
            runPropsBold.Append(New FontSize() With {.Val = fontSize})
        End If
        runBold.Append(runPropsBold)
        
        ' Construire le texte : Nom Prénom dit Pseudonyme
        Dim nom As String = If(info.Nom, "").ToUpper()
        Dim prenom As String = If(info.Prenom, "")
        Dim ligne1 As String = $"{nom} {prenom}".Trim()
        
        If Not String.IsNullOrEmpty(info.Pseudonyme) Then
            ligne1 &= $" dit {info.Pseudonyme}"
        End If
        
        Dim text1 As New Text(ligne1)
        text1.Space = SpaceProcessingModeValues.Preserve
        runBold.Append(text1)
        paragraph.Append(runBold)
        
        ' === SAUT DE LIGNE ===
        Dim runBreak As New Run()
        runBreak.Append(New Break())
        paragraph.Append(runBreak)
        
        ' === LIGNE 2 : Rôle - NORMAL ===
        Dim runNormal As New Run()
        ' Appliquer la police si spécifiée
        If Not String.IsNullOrEmpty(fontName) OrElse Not String.IsNullOrEmpty(fontSize) Then
            Dim runPropsNormal As New RunProperties()
            If Not String.IsNullOrEmpty(fontName) Then
                runPropsNormal.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
            End If
            If Not String.IsNullOrEmpty(fontSize) Then
                runPropsNormal.Append(New FontSize() With {.Val = fontSize})
            End If
            runNormal.Append(runPropsNormal)
        End If
        
        Dim ligne2 As String = ConvertRoleGenre(info.Role, info.Genre)
        Dim text2 As New Text(ligne2)
        text2.Space = SpaceProcessingModeValues.Preserve
        runNormal.Append(text2)
        paragraph.Append(runNormal)
        
        cell.Append(paragraph)
        Return cell
    End Function
    
    ''' <summary>
    ''' Crée une cellule de tableau avec police personnalisée
    ''' </summary>
    Private Function CreateCellWithFont(text As String, isBold As Boolean, isCenter As Boolean, fontName As String, fontSize As String) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' Alignement
        If isCenter Then
            Dim paragraphProperties As New ParagraphProperties()
            paragraphProperties.Append(New Justification() With {.Val = JustificationValues.Center})
            paragraph.Append(paragraphProperties)
        End If

        ' Run
        Dim run As New Run()
        Dim runProperties As New RunProperties()
        
        If isBold Then
            runProperties.Append(New Bold())
        End If
        
        ' Appliquer la police si spécifiée
        If Not String.IsNullOrEmpty(fontName) Then
            runProperties.Append(New RunFonts() With {.Ascii = fontName, .HighAnsi = fontName})
        End If
        If Not String.IsNullOrEmpty(fontSize) Then
            runProperties.Append(New FontSize() With {.Val = fontSize})
        End If
        
        If runProperties.HasChildren Then
            run.Append(runProperties)
        End If
        
        run.Append(New Text(text))
        
        paragraph.Append(run)
        cell.Append(paragraph)

        Return cell
    End Function
    
    ''' <summary>
    ''' Convertit le rôle avec genre (Auteur, Compositeur, etc.) - SANS "d'" ou "de"
    ''' </summary>
    Private Function ConvertRoleGenre(role As String, genre As String) As String
        If String.IsNullOrEmpty(role) Then Return ""
        
        If genre = "MME" Then
            Select Case role.ToUpper()
                Case "A" : Return "Autrice"
                Case "C" : Return "Compositrice"
                Case "AR" : Return "Arrangeuse"
                Case "AD" : Return "Adaptatrice"
                Case "AC" : Return "Autrice-Compositrice"
                Case Else : Return role
            End Select
        Else
            Select Case role.ToUpper()
                Case "A" : Return "Auteur"
                Case "C" : Return "Compositeur"
                Case "AR" : Return "Arrangeur"
                Case "AD" : Return "Adaptateur"
                Case "AC" : Return "Auteur-Compositeur"
                Case Else : Return role
            End Select
        End If
    End Function
    
    ''' <summary>
    ''' Crée une cellule avec contenu multiligne
    ''' </summary>
    Private Function CreateCellMultiline(content As String, Optional isBold As Boolean = False, Optional isCenter As Boolean = False) As TableCell
        Dim cell As New TableCell()

        ' Propriétés de cellule
        Dim cellProperties As New TableCellProperties()
        Dim cellBorders As New TableCellBorders()
        cellBorders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellBorders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4})
        cellProperties.Append(cellBorders)
        cell.Append(cellProperties)

        ' Paragraphe
        Dim paragraph As New Paragraph()
        
        ' Alignement
        If isCenter Then
            Dim paragraphProperties As New ParagraphProperties()
            paragraphProperties.Append(New Justification() With {.Val = JustificationValues.Center})
            paragraph.Append(paragraphProperties)
        End If

        ' Run avec gestion des sauts de ligne
        Dim run As New Run()
        If isBold Then
            Dim runProperties As New RunProperties()
            runProperties.Append(New Bold())
            run.Append(runProperties)
        End If
        
        ' Gérer les sauts de ligne
        Dim lines As String() = content.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
        For i As Integer = 0 To lines.Length - 1
            If i > 0 Then
                run.Append(New Break())
            End If
            Dim textElement As New Text(lines(i))
            textElement.Space = SpaceProcessingModeValues.Preserve
            run.Append(textElement)
        Next
        
        paragraph.Append(run)
        cell.Append(paragraph)

        Return cell
    End Function

    ''' <summary>
    ''' Info de signature
    ''' </summary>
    Private Class SignatureInfo
        Public Property Designation As String
        Public Property Role As String
        Public Property RoleCode As String
        Public Property Genre As String
    End Class
    
    ''' <summary>
    ''' Info de créateur pour les tableaux de répartition
    ''' </summary>
    Private Class CreatorInfo
        Public Property Nom As String
        Public Property Prenom As String
        Public Property Pseudonyme As String
        Public Property Genre As String
        Public Property Role As String
        Public Property PH As Double
    End Class
    
    ''' <summary>
    ''' Vérifie si un ayant droit est membre SACEM
    ''' Retourne True si SACEM ou si SocieteGestion est vide (défaut = SACEM)
    ''' Retourne False si membre d'une autre société (GEMA, KODA, etc.)
    ''' </summary>
    Private Function IsSACEM(ayant As AyantDroit) As Boolean
        Dim societe As String = If(ayant.Identite.SocieteGestion, "").Trim().ToUpper()
        Return String.IsNullOrEmpty(societe) OrElse societe = "SACEM"
    End Function
End Class
