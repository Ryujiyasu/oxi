Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value(SizeConst()) As Long
Probe = CStr(UBound(value))
End Function
Private Function SizeConst() As Long
SizeConst = 2
End Function
