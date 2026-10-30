Option Explicit
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Function Inside() As String
 Inside = Me.Name & ":" & Me.CodeName & ":" & CStr(Me Is ThisWorkbook) & ":" & CStr(Me.Worksheets.Count)
End Function
