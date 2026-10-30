Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(VBA.CLng(-1)) As Long
Probe = CStr(Err.Number) & ":ran"
End Function
