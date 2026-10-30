Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(Round(1.5)) As Long
Probe = CStr(UBound(value))
End Function
