Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim value As Variant
On Error Resume Next
value = Fix(Nothing)
Probe = CStr(Err.Number) & ":" & TypeName(value)
If Not IsObject(value) Then Probe = Probe & ":" & CStr(value)
End Function
