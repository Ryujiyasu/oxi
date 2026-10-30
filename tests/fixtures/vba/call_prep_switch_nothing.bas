Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value As Variant, result As Variant
Set value = Nothing
On Error Resume Next
result = Switch(Nothing, 1, True, 2)
Probe = CStr(Err.Number) & ":" & TypeName(result)
End Function
