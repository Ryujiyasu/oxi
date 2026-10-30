Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = CStr(CLng(2))
End Function
Private Function CLng(ByVal number As Long) As Long
CLng = number
End Function
