Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public trace As String
Public Function Probe() As String
number = 1
Change number, number
Probe = trace & "|" & number
End Function
Private Sub Change(ParamArray values() As Variant)
values(0) = values(0) + 1
trace = values(0) & ":" & values(1) & ":" & number
values(1) = values(1) + 10
End Sub
