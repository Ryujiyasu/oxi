Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const first = second
Const second = 2
Probe = CStr(Err.Number) & ":ran"
End Function
