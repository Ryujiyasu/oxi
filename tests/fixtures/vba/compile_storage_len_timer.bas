Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String

On Error Resume Next
Dim result As Variant
result = Len(Timer)
Probe = Err.Number & ":" & TypeName(result)
If Not IsNull(result) Then Probe = Probe & ":" & CStr(result)
End Function
