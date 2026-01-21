Imports System.Diagnostics
Imports System.IO

''' <summary>
''' Service pour convertir les fichiers Excel en JSON via le script Python
''' </summary>
Public Class ExcelToJsonConverter
    Private ReadOnly _scriptPath As String
    Private ReadOnly _pythonPath As String
    
    ''' <summary>
    ''' Crée une instance du convertisseur
    ''' </summary>
    ''' <param name="scriptPath">Chemin vers Xlstojson.py (optionnel, par défaut Scripts/Xlstojson.py)</param>
    ''' <param name="pythonPath">Chemin vers l'exécutable Python (optionnel, par défaut "python")</param>
    Public Sub New(Optional scriptPath As String = Nothing, Optional pythonPath As String = Nothing)
        ' Chemin par défaut du script
        If String.IsNullOrEmpty(scriptPath) Then
            _scriptPath = Path.Combine(Application.StartupPath, "Scripts", "Xlstojson", "Xlstojson.py")
        Else
            _scriptPath = scriptPath
        End If
        
        ' Chemin Python par défaut
        _pythonPath = If(String.IsNullOrEmpty(pythonPath), "python", pythonPath)
    End Sub
    
    ''' <summary>
    ''' Convertit un fichier Excel en JSON
    ''' </summary>
    ''' <param name="xlsxPath">Chemin du fichier Excel</param>
    ''' <param name="outputDirectory">Dossier de sortie pour le JSON (optionnel)</param>
    ''' <returns>Chemin du fichier JSON généré, ou Nothing en cas d'erreur</returns>
    Public Function ConvertFile(xlsxPath As String, Optional outputDirectory As String = Nothing) As String
        Try
            If Not File.Exists(xlsxPath) Then
                Throw New FileNotFoundException($"Fichier Excel non trouvé: {xlsxPath}")
            End If
            
            If Not File.Exists(_scriptPath) Then
                Throw New FileNotFoundException($"Script Python non trouvé: {_scriptPath}")
            End If
            
            ' Dossier de sortie par défaut = même dossier que le fichier Excel
            If String.IsNullOrEmpty(outputDirectory) Then
                outputDirectory = Path.GetDirectoryName(xlsxPath)
            End If
            
            ' Créer le dossier de sortie si nécessaire
            If Not Directory.Exists(outputDirectory) Then
                Directory.CreateDirectory(outputDirectory)
            End If
            
            ' Créer un dossier temporaire pour le fichier Excel
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "SACEMConverter_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempDir)
            
            ' Copier le fichier Excel dans le dossier temporaire
            Dim tempXlsx As String = Path.Combine(tempDir, Path.GetFileName(xlsxPath))
            File.Copy(xlsxPath, tempXlsx, True)
            
            ' Exécuter le script Python
            Dim psi As New ProcessStartInfo()
            psi.FileName = _pythonPath
            psi.Arguments = $"""{_scriptPath}"" ""{tempDir}"" ""{outputDirectory}"""
            psi.UseShellExecute = False
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True
            psi.WorkingDirectory = Path.GetDirectoryName(_scriptPath)
            
            Using process As Process = Process.Start(psi)
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim errors As String = process.StandardError.ReadToEnd()
                process.WaitForExit()
                
                ' Nettoyer le dossier temporaire
                Try
                    Directory.Delete(tempDir, True)
                Catch
                End Try
                
                If process.ExitCode <> 0 Then
                    Throw New Exception($"Erreur Python: {errors}")
                End If
                
                ' Trouver le fichier JSON généré
                Dim jsonFiles = Directory.GetFiles(outputDirectory, "*.json", SearchOption.TopDirectoryOnly)
                Dim baseName As String = Path.GetFileNameWithoutExtension(xlsxPath)
                
                For Each jsonFile In jsonFiles
                    If Path.GetFileNameWithoutExtension(jsonFile).Contains(baseName) OrElse
                       jsonFile.EndsWith("_.json") Then
                        Return jsonFile
                    End If
                Next
                
                ' Retourner le dernier fichier JSON créé
                If jsonFiles.Length > 0 Then
                    Return jsonFiles.OrderByDescending(Function(f) File.GetCreationTime(f)).First()
                End If
            End Using
            
            Return Nothing
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur conversion Excel->JSON: {ex.Message}")
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' Convertit tous les fichiers Excel d'un dossier en JSON
    ''' </summary>
    ''' <param name="xlsxDirectory">Dossier contenant les fichiers Excel</param>
    ''' <param name="jsonDirectory">Dossier de sortie pour les JSON</param>
    ''' <returns>Liste des fichiers JSON générés</returns>
    Public Function ConvertDirectory(xlsxDirectory As String, jsonDirectory As String) As List(Of String)
        Try
            If Not Directory.Exists(xlsxDirectory) Then
                Throw New DirectoryNotFoundException($"Dossier Excel non trouvé: {xlsxDirectory}")
            End If
            
            If Not File.Exists(_scriptPath) Then
                Throw New FileNotFoundException($"Script Python non trouvé: {_scriptPath}")
            End If
            
            ' Créer le dossier de sortie si nécessaire
            If Not Directory.Exists(jsonDirectory) Then
                Directory.CreateDirectory(jsonDirectory)
            End If
            
            ' Exécuter le script Python
            Dim psi As New ProcessStartInfo()
            psi.FileName = _pythonPath
            psi.Arguments = $"""{_scriptPath}"" ""{xlsxDirectory}"" ""{jsonDirectory}"""
            psi.UseShellExecute = False
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True
            psi.WorkingDirectory = Path.GetDirectoryName(_scriptPath)
            
            Dim generatedFiles As New List(Of String)
            
            Using process As Process = Process.Start(psi)
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim errors As String = process.StandardError.ReadToEnd()
                process.WaitForExit()
                
                If process.ExitCode <> 0 Then
                    Throw New Exception($"Erreur Python: {errors}")
                End If
                
                ' Lister tous les fichiers JSON du dossier de sortie
                generatedFiles.AddRange(Directory.GetFiles(jsonDirectory, "*.json"))
            End Using
            
            Return generatedFiles
            
        Catch ex As Exception
            Debug.WriteLine($"Erreur conversion dossier Excel->JSON: {ex.Message}")
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' Vérifie si Python est disponible
    ''' </summary>
    Public Function IsPythonAvailable() As Boolean
        Try
            Dim psi As New ProcessStartInfo()
            psi.FileName = _pythonPath
            psi.Arguments = "--version"
            psi.UseShellExecute = False
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            psi.CreateNoWindow = True
            
            Using process As Process = Process.Start(psi)
                process.WaitForExit(5000)
                Return process.ExitCode = 0
            End Using
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Vérifie si le script Python existe
    ''' </summary>
    Public Function IsScriptAvailable() As Boolean
        Return File.Exists(_scriptPath)
    End Function
    
    ''' <summary>
    ''' Retourne le chemin du script Python
    ''' </summary>
    Public ReadOnly Property ScriptPath As String
        Get
            Return _scriptPath
        End Get
    End Property
End Class
