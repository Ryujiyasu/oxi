Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As String = 123
Probe = CStr(VarType(value)) & ":" & CStr(value)
End Function
