Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
values(1) = values(1) + 1
Probe = CStr(values(1))
End Function
Attribute VB_Name = "CounterModule"
Public values(1 To 2) As Long
