Attribute VB_Name = "DriverModule"
Option Explicit
Public number As String * 4
Public trace As String
Public Function Probe() As String
number = "abcd"
Inner number
Probe = trace & "|" & number
End Function
Private Sub Inner(ByRef value As Variant)
On Error Resume Next
value = 2.6
trace = TypeName(value) & ":" & value & ":" & number
value = "bad"
trace = trace & ":" & Err.Number
End Sub
