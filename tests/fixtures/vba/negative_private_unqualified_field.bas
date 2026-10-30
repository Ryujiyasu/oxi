Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
If False Then Probe = CStr(hiddenValue)
Probe = "after"
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Private hiddenValue As Long
