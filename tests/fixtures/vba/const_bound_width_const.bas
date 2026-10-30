Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const size = 3
Dim value As String * size
Probe = CStr(Len(value))
End Function
