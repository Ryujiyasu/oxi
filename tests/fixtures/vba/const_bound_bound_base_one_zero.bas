Attribute VB_Name = "DriverModule"
Option Explicit
Option Base 1
Public Function Probe() As String
On Error Resume Next
Dim value(0) As Long
Probe = CStr(LBound(value)) & ":" & CStr(UBound(value))
End Function
