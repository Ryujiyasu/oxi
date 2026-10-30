Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value() As Long
ReDim value(1 To 0)
Probe = CStr(Err.Number) & ":ran"
End Function
