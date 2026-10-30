Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public trace As String
Public Function Probe() As String
number = 1
Change number
Probe = trace & "|" & number
End Function
Private Sub Change(ParamArray values() As Variant)
On Error Resume Next
values(0) = 2.6
trace = TypeName(number) & ":" & number & ":" & TypeName(values(0))
values(0) = "bad"
trace = trace & ":" & Err.Number
End Sub
