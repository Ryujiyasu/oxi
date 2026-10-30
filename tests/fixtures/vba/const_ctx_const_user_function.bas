Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const size = SizeConst()
Probe = CStr(Err.Number) & ":ran"
End Function
Private Function SizeConst() As Long
SizeConst = 2
End Function
