Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Long
Probe = CStr(Err.Number) & ":ran"
End Function
