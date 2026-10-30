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
Dim item As Variant
Dim total As Long
values(0) = values(0) + 1
For Each item In values
total = total + item
Next item
trace = Join(values, ",") & ":" & total
End Sub
