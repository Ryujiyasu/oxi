Option Explicit
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Function Inside() As String
 Range("A1").Value = 7
 Inside = Name & ":" & CodeName & ":" & CStr(Range("A1").Value)
End Function
