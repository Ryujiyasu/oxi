Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
LeftModule.items(1) = LeftModule.items(1) + 1
RightModule.items(1) = RightModule.items(1) + 10
Probe = LeftModule.items(1) & "|" & RightModule.items(1)
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public items(1 To 1) As Long

Attribute VB_Name = "RightModule"
Option Explicit
Public items(1 To 1) As Long
