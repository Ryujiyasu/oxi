Option Explicit
Private count As Long
Public Function Report() As String
 count = count + 1
 Report = Me.Name & ":" & Me.CodeName & ":" & TypeName(Me) & ":" & CStr(Me.Index) & ":" & CStr(count)
End Function
