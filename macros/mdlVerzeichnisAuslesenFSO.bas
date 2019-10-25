Attribute VB_Name = "mdlVerzeichnisAuslesenFSO"
Option Explicit
Private mlngRow As Long

Private Enum FilePropertiesEnum
    udPath = 1&
    udName = 2&
    udFileLen = 4&
    udDateLastModified = 8&
    udDateLastAccessed = 16&
    udDateCreated = 32&
    udFileType = 64&
    udFileAttr = 128&
    udShortPath = 256&
End Enum

Public Sub ScanDir()
Dim strPath As String
Dim objWsh As Excel.Worksheet
Dim strFilter As String
Dim lngCol As Long

strPath = "C:\Users\Daniel\Documents\"
Set objWsh = ThisWorkbook.ActiveSheet
strFilter = ""
lngCol = 1
mlngRow = 1

Call getFilesInDir( _
                   strPath:=strPath, _
                    objWsh:=objWsh, _
                 strFilter:=strFilter, _
                    lngCol:=lngCol, _
            FileProperties:=udPath Or udName Or udFileLen Or udDateLastModified Or udDateLastAccessed Or udDateCreated Or udFileType Or udFileAttr Or udShortPath)
End Sub

Private Sub getFilesInDir( _
                          ByVal strPath As String, _
                          ByRef objWsh As Excel.Worksheet, _
                          Optional ByVal strFilter As String = "", _
                          Optional ByVal lngCol As Long = 1, _
                          Optional ByVal FileProperties As FilePropertiesEnum = udPath)
                          
Dim objFSO As Scripting.FileSystemObject
Dim objDir As Scripting.Folder
Dim objFiles As Scripting.Files
Dim objFile As Scripting.File
Dim objSub As Scripting.Folder
Dim i As Integer

On Error Resume Next

If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
strFilter = "*" & strFilter & "*"

'Instanz der Klasse FileSystemObject erzeugen
Set objFSO = New Scripting.FileSystemObject

'Verzeichnisliste des aktuellen Ordners
Set objDir = objFSO.GetFolder(strPath)

'Liste aller Dateien in dem Ordner
Set objFiles = objDir.Files

    For Each objFile In objFiles
        
        If LCase(objFile.Name) Like LCase(strFilter) Then
        
            i = 1
            '**************** Prüfung der Flags ****************
            'Dateipfad ausgeben
            If CBool(FileProperties And udPath) Then
                objWsh.Cells(mlngRow, lngCol).Value = objFile.Path
            End If
            
            'Dateiname ausgeben
            If CBool(FileProperties And udName) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.Name
                i = i + 1
            End If
            
            'Dateigröße in Bytes ausgeben
            If CBool(FileProperties And udFileLen) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.Size
                i = i + 1
            End If
            
            'Letzte Änderung
            If CBool(FileProperties And udDateLastModified) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.DateLastModified
                i = i + 1
            End If
            
            'Letztes Öffnen
            If CBool(FileProperties And udDateLastAccessed) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.DateLastAccessed
                i = i + 1
            End If
            
            'Erstellungsdatum
            If CBool(FileProperties And udDateCreated) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.DateCreated
                i = i + 1
            End If
            
            'Dateityp ausgeben
            If CBool(FileProperties And udFileType) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.Type
                i = i + 1
            End If
            
            'Dateiattribut ausgeben
            If CBool(FileProperties And udFileAttr) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.Attributes
                i = i + 1
            End If
            
            'Shortpath ausgeben
            If CBool(FileProperties And udShortPath) Then
                objWsh.Cells(mlngRow, lngCol + i).Value = objFile.ShortPath
                i = i + 1
            End If
            
            mlngRow = mlngRow + 1
            
        End If
    
    Next objFile
    
    'Unterverzeichnisse abarbeiten
    For Each objSub In objDir.SubFolders
    
        Call getFilesInDir( _
                           strPath:=objSub.Path, _
                            objWsh:=objWsh, _
                         strFilter:=strFilter, _
                            lngCol:=lngCol, _
                    FileProperties:=FileProperties)
                           
    Next objSub

End Sub
