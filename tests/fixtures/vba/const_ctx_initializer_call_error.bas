Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = CByte(1 / 0)
Probe = CStr(value)
End Function
