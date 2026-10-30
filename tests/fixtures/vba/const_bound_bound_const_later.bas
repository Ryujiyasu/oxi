Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(size) As Long
Const size = 2
Probe = CStr(UBound(value))
End Function
