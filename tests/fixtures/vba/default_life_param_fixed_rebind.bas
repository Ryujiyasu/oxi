Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value As String * 4
value = "abcd"
Change value
Probe = "[" & value & "]"
End Function
Private Sub Change(ParamArray values() As Variant)
values(0) = "xyz"
values = Array("123456")
End Sub
