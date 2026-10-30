Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
If False Then Probe = CStr(LeftModule.Hidden())
Probe = "after"
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Private Function Hidden() As Long
Hidden = 1
End Function
