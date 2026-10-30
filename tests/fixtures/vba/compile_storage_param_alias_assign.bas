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
values(0) = 2
values = Array(3, 4)
trace = values(0) & ":" & values(1) & ":" & number
values(0) = 5
trace = trace & ":" & number
End Sub
