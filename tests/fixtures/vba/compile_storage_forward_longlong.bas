Attribute VB_Name = "DriverModule"
Option Explicit
Public number As LongLong
Public trace As String
Public Function Probe() As String
number = 1
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
