Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Long = 2147483647.5#
Probe = CStr(VarType(value)) & ":" & CStr(value)
End Function
