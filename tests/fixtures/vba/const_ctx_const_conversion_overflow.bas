Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const size = CByte(300)
Probe = CStr(Err.Number) & ":ran"
End Function
