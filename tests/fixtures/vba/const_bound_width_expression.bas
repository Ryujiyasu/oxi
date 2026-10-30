Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value As String * (2 + 1)
Probe = CStr(Len(value))
End Function
