Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value("bad") As Long
Probe = CStr(LBound(value)) & ":" & CStr(UBound(value))
End Function
