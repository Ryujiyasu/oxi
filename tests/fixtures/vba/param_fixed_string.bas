Attribute VB_Name = "DriverModule"
Option Explicit
Public text As String * 4
Public trace As String
Public Function Probe() As String
text = "abcd"
Change text, text
Probe = trace & "|" & text
End Function
Private Sub Change(ParamArray values() As Variant)
values(0) = "longtext"
trace = values(0) & ":" & values(1) & ":" & text
End Sub
