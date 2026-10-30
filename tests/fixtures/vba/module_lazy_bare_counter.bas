Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
count = count + 1
Probe = CStr(count)
End Function
Attribute VB_Name = "CounterModule"
Public count As Long
