Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Integer = 1.5
Probe = CStr(value)
End Function
