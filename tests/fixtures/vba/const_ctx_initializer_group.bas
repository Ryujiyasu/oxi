Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = (2 + 3)
Probe = CStr(value)
End Function
