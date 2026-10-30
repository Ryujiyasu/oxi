Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Boolean = Null
Probe = CStr(VarType(value)) & ":" & CStr(value)
End Function
