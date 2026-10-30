Option Explicit
Public Function Helper(ByVal required As Long) As Long
 Helper = required
End Function
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
Public Function Inside() As String
 Me.Range("A1").Value = 7
 Inside = Me.Name & ":" & Me.CodeName & ":" & CStr(Me.Range("A1").Value) & ":" & CStr(Me Is Worksheets(1)) & ":" & CStr(Me.Parent Is ThisWorkbook)
End Function
