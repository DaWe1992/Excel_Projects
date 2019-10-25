Attribute VB_Name = "mdlVerzeichnisAuslesen"
Option Explicit
Private mlngRow As Long

Public Sub GetFileNamesInDirectoryExecute()
Dim strPath As String
Dim objWS As Worksheet
Dim strFilter As String

strPath = InputBox(prompt:="Bitte geben Sie ein Verzeichnis ein:" & vbCr & vbCr & _
                "Bitte beachten Sie, dass die Auswertung u. U. sehr lange dauern kann!", _
                    Title:="Verzeichnisauswahl", _
                  Default:="C:\Users\Daniel\Documents\")
                  
'Überprüfung von strPath
If strPath = "" Then Exit Sub
If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"

Set objWS = ThisWorkbook.ActiveSheet
objWS.Cells.ClearContents

strFilter = ""

mlngRow = 2 'Start bei angegebener Zeile

Application.ScreenUpdating = False

'Prozeduraufruf
Call getFileNamesInDirectory( _
                             strFolder:=strPath, _
                                 objWS:=objWS, _
                             strFilter:=strFilter, _
                            blnFileLen:=True, _
                       blnFileDateTime:=True)
                             
Application.ScreenUpdating = True
                             
End Sub

Private Sub getFileNamesInDirectory( _
                                    ByRef strFolder As String, _
                                    ByRef objWS As Worksheet, _
                                    Optional ByRef strFilter As String = "", _
                                    Optional ByVal blnFileLen As Boolean = False, _
                                    Optional ByVal blnFileDateTime As Boolean = False)
Dim strFile As String
Dim i As Integer
Dim astrFolders() As String

On Error Resume Next

ReDim astrFolders(1 To 20)

'Filterkriterium einkleiden
strFilter = "*" & strFilter & "*"

i = 0
'Dateien mit sämltichen Attributen durchlaufen
strFile = Dir$(strFolder, vbNormal Or vbHidden Or _
                          vbDirectory Or vbArchive Or vbSystem Or _
                          vbReadOnly Or vbVolume)

    Do While strFile <> ""
    
        'Dateiattribute auf vbDirectory prüfen
        If GetAttr(strFolder & strFile) And vbDirectory Then
            
            'Nur Unterverzeichnisse aufnehmen
            'Ein Punkt steht für ein übergeordnetes Verzeichnis
            If Right$(strFile, 1) <> "." Then
    
                i = i + 1
                If i > UBound(astrFolders) Then
                    
                    'Array anpassen, wenn Index nicht ausreicht
                    ReDim Preserve astrFolders(1 To (i + 1))
                    
                End If
                
                'Ordner in Array speichern
                astrFolders(i) = strFile & "\"
                    
            End If
            
        Else
        
            'Nur ausgeben, wenn Datei den Filterkriterien entspricht
            If LCase(strFile) Like LCase(strFilter) Then
        
                objWS.Cells(mlngRow, 1).Value = strFolder & strFile
            
                'Größe der Datei in Bytes ausgeben
                If blnFileLen Then
                    objWS.Cells(mlngRow, 2).Value = _
                    FileLen(strFolder & strFile)
                End If
            
                'Ausgabe der letzten Änderung
                If blnFileDateTime Then
                    objWS.Cells(mlngRow, 3).Value = _
                    FileDateTime(strFolder & strFile)
                End If
            
                mlngRow = mlngRow + 1
            
            End If
            
        End If
        
        strFile = Dir$()
        
    Loop
    
    If i = 0 Then Exit Sub
    
    ReDim Preserve astrFolders(1 To i)
    
    'Untergeordnete Verzeichnisse durchlaufen
    For i = 1 To UBound(astrFolders)
        
        'Rekursion
        Call getFileNamesInDirectory( _
                                     strFolder & astrFolders(i), _
                                     objWS, _
                                     strFilter, _
                                     blnFileLen, _
                                     blnFileDateTime)
                            
    Next i
                            
End Sub
