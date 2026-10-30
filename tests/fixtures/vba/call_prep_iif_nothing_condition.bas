Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim value As Variant
On Error Resume Next
value = IIf(Nothing, 1, 2)
Probe = CStr(Err.Number)
End Function
