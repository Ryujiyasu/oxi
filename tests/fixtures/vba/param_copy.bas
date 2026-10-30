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
Dim copied As Variant
copied = values
values(0) = values(0) + 1
trace = copied(0) & ":" & values(0) & ":" & number
copied(0) = copied(0) + 10
trace = trace & ":" & values(0) & ":" & number
End Sub
