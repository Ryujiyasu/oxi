Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Date = 1E+15
Probe = CStr(VarType(value)) & ":" & CStr(value)
End Function
