Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public trace As String
Public Function Probe() As String
number = 1
Change number, (number), number + 0
Probe = trace & "|" & number
End Function
Private Sub Change(ParamArray values() As Variant)
values(0) = values(0) + 1
values(1) = values(1) + 10
values(2) = values(2) + 100
trace = Join(values, ",")
End Sub
