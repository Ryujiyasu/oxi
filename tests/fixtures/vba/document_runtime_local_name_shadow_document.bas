Option Explicit
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Function Inside() As String
 Dim Name As String
 Name = "local"
 Inside = Name & ":" & Me.Name
End Function
