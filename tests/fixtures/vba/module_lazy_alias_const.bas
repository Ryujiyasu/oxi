Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Probe = CStr(CounterModule.High)
End Function
Attribute VB_Name = "CounterModule"
Private Const seed = 2
Public Const High = seed + 1
