Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
CounterModule.count = CounterModule.count + 1
Probe = CStr(CounterModule.count)
End Function
Attribute VB_Name = "CounterModule"
Public count As Long
