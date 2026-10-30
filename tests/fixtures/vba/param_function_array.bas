Attribute VB_Name = "DriverModule"
Option Explicit
Public calls As Long
Public Function MakeValues() As Variant
calls = calls + 1
MakeValues = Array(1, 2)
End Function
Public Function Probe() As String
Change MakeValues()(1)
Probe = CStr(calls)
End Function
Private Sub Change(ByRef value As Variant)
value = value + 10
End Sub
