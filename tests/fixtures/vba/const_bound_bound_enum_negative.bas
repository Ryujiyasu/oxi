Attribute VB_Name = "DriverModule"
Option Explicit
Private Enum Limits
Low = -1
High = 2
End Enum
Public Function Probe() As String
On Error Resume Next
Dim value(Limits.Low) As Long
Probe = CStr(LBound(value)) & ":" & CStr(UBound(value))
End Function
