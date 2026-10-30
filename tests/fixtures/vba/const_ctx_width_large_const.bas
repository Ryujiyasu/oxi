Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const size = 65536
Dim value As String * size
Probe = CStr(Err.Number) & ":ran"
End Function
