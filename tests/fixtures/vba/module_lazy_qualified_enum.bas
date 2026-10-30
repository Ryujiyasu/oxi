Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Probe = CStr(Bounds.TopValue)
End Function
Attribute VB_Name = "CounterModule"
Public Enum Bounds
TopValue = 2
End Enum
