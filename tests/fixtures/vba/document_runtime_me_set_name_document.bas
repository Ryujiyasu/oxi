Option Explicit
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Function Inside() As String
 Me.Name = "Changed"
 Inside = Me.Name & ":" & Me.CodeName
End Function
