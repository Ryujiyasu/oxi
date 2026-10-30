Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim value As Variant
On Error Resume Next
value = Nothing & "x"
Probe = CStr(Err.Number) & ":" & TypeName(value)
End Function
