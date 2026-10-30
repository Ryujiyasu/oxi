Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value As Variant
value = 123
On Error Resume Next
Dim result As Variant
result = Len(value + 1)
Probe = Err.Number & ":" & TypeName(result)
If Not IsNull(result) Then Probe = Probe & ":" & CStr(result)
End Function
