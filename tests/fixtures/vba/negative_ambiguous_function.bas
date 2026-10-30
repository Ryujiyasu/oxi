Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
If False Then Probe = CStr(Clash())
Probe = "after"
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public Function Clash() As Long
Clash = 1
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Public Function Clash() As Long
Clash = 1
End Function
