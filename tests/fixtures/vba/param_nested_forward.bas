Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public trace As String
Public Function Probe() As String
number = 1
Outer number
Probe = trace & "|" & number
End Function
Private Sub Outer(ParamArray values() As Variant)
values(0) = values(0) + 1
Inner values(0)
trace = values(0) & ":" & number
End Sub
Private Sub Inner(ParamArray values() As Variant)
values(0) = values(0) + 10
End Sub
