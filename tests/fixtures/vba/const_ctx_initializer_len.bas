Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = Len("ab")
Probe = CStr(value)
End Function
