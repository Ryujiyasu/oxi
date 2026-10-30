Attribute VB_Name = "DriverModule"
Option Explicit
Public values As Variant
Public Function Probe() As String
values = Array(Array(1, 2))
Change values(0)(1)
Probe = CStr(values(0)(1))
End Function
Private Sub Change(ByRef value As Variant)
value = value + 10
End Sub
