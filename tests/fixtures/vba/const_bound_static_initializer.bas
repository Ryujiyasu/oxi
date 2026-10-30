Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Static value As Long = 7
Probe = CStr(Err.Number) & ":ran"
End Function
