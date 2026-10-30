Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = VBA.CLng(2)
Probe = CStr(value)
End Function
