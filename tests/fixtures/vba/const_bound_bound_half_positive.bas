Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(1.5) As Long
Probe = CStr(LBound(value)) & ":" & CStr(UBound(value))
End Function
