Option Explicit
Public Function Inside() As String
 Range("A1").Value = 7
 Inside = CStr(Me.Range("A1").Value) & ":" & CStr(Worksheets(2).Range("A1").Value)
End Function
