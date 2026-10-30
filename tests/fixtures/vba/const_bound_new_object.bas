Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value As New Object
Probe = CStr(Err.Number) & ":ran"
End Function
