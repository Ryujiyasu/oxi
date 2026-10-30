Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = Abs(-1)
Probe = CStr(value)
End Function
