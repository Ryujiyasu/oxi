Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
LeftModule.held.amount = LeftModule.held.amount + 1
Probe = CStr(LeftModule.held.amount)
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public Type Entry
amount As Long
End Type
Public held As Entry
