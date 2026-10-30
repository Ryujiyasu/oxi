Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Bump sharedValue
RightModule.Bump n:=LeftModule.sharedValue
Probe = sharedValue & "|" & LeftModule.sharedValue & "|" & CStr(Len(sharedValue))
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public sharedValue As Long

Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Bump(ByRef n As Long)
n = n + 1
End Sub
