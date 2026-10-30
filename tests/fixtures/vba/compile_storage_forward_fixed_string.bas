Attribute VB_Name = "DriverModule"
Option Explicit
Public number As String * 4
Public trace As String
Public Function Probe() As String
number = "abcd"
Outer number
Probe = trace & "|" & TypeName(number) & ":" & CStr(number)
End Function
Private Sub Outer(ParamArray values() As Variant)
Inner values(0)
End Sub
Private Sub Inner(ByRef value As Variant)
On Error Resume Next
value = 2.6
trace = TypeName(number) & ":" & CStr(number)
value = "bad"
trace = trace & ":" & TypeName(number) & ":" & Err.Number
End Sub
