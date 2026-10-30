Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Const value = Nothing
On Error Resume Next
Probe = CStr(Err.Number)
End Function
