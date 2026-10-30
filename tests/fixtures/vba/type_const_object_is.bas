Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = (Nothing Is Nothing)
Probe = CStr(value)
End Function
