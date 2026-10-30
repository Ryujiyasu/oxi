Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value As Variant
On Error Resume Next
value = VBA.CVar(Nothing)
Probe = CStr(Err.Number) & ":" & TypeName(value)
End Function
