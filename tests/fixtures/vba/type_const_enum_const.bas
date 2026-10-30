Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Choice
One = 1
End Enum
Public Function Probe() As String
On Error Resume Next
Const value As Choice = 2
Probe = CStr(VarType(value)) & ":" & CStr(value)
End Function
