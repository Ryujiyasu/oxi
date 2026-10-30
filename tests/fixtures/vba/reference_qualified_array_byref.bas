Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
LeftModule.items(1) = LeftModule.items(1) + 1
RightModule.Bump LeftModule.items(1)
Probe = CStr(LeftModule.items(1))
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public items(1 To 1) As Long

Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Bump(ByRef value As Long)
value = value + 1
End Sub
