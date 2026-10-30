Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(1 \ 0) As Long
Probe = CStr(UBound(value))
End Function
