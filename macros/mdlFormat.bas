Attribute VB_Name = "mdlFormat"
Option Explicit

Private Enum FormatTextEnum
    UpperCase = 1&
    LowerCase = 2&
    TrimSpace = 4&
    TrimTab = 8&
    Reverse = 16&
    AddPeriod = 32&
End Enum

Private Function FormatText( _
                            ByVal strText As String, _
                            ByVal udFormat As FormatTextEnum _
                            ) As String
                                                
    If CBool(udFormat And UpperCase) Then
        strText = UCase(strText)
    End If
    
    If CBool(udFormat And LowerCase) Then
        strText = LCase(strText)
    End If
    
    If CBool(udFormat And TrimSpace) Then
        strText = Replace(strText, " ", "")
    End If
    
    If CBool(udFormat And TrimTab) Then
        strText = Replace(strText, vbTab, "")
    End If
    
    If CBool(udFormat And Reverse) Then
        strText = StrReverse(strText)
    End If
    
    If CBool(udFormat And AddPeriod) Then
        strText = strText & "."
    End If
              
    FormatText = strText
    
End Function

Public Sub Test()
    MsgBox _
          prompt:=FormatText( _
                             strText:="Das ist ein Beispieltext", _
                            udFormat:=UpperCase Or AddPeriod), _
         Buttons:=vbOKOnly, _
           Title:="Ergebnis"
End Sub
                            

