Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Byte = 255.4
Probe = CStr(VarType(value)) & ":" & CStr(value)
End Function
