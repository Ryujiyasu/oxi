Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value As String * 3#
Probe = CStr(Err.Number) & ":ran"
End Function
