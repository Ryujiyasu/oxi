Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim LongPtr As Long
LongPtr = 7
Probe = CStr(LongPtr)
End Function
