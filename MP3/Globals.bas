Attribute VB_Name = "Globals"
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public gbPlaylist As Boolean


Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub


Public Function GetId3(Filename As String)
Dim psTemp As String * 3
    Open Filename For Binary As #1
    Get #1, FileLen(Filename) - 127, psTemp
    
    If psTemp = "psTemp" Then
        Get #1, FileLen(Filename) - 124, id3Info
    Else
        MsgBox "This Mp3 Does Not Contain an ID3 psTemp"
    End If
    
    Close #1
End Function
