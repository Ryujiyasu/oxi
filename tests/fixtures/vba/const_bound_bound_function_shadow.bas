Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(CLng(2)) As Long
Probe = CStr(UBound(value))
End Function
Private Function CLng(ByVal number As Long) As Long
CLng = number
End Function
