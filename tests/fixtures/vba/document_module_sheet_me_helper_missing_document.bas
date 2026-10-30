Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Sub Inside()
 Dim n As Long
 n = Me.Helper()
End Sub
