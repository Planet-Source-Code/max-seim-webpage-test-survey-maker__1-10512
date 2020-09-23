Attribute VB_Name = "testmaker"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function FileExists(strfile As String) As Boolean
If Dir(strfile) <> "" Then
FileExists = True
Else
FileExists = False
End If
End Function
Sub ErrHandler()
    ErrDesc = Err.Description
    ErrNum = Err.Number
    Beep
    MsgBox "Error number " & ErrNum & " has occured because: " & _
    ErrDesc, vbCritical, "Error"
    Exit Sub
    
End Sub




