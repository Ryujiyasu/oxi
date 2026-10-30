Attribute VB_Name = "DriverModule"
Option Explicit
Private Enum Limits
Low = 1
High
End Enum
Public Function Probe() As String
On Error Resume Next
Dim value(Limits.High) As Long
Probe = CStr(UBound(value))
End Function
