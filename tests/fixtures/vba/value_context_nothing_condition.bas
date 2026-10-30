Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
If Nothing Then
Probe = "taken"
End If
Probe = CStr(Err.Number)
End Function
