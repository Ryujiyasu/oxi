Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = LeftModule.StepLeft() & "|" & RightModule.StepRight()
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Private stored As Long
Private Function Helper() As Long
Static count As Long
count = count + 1
Helper = count
End Function
Public Function StepLeft() As String
stored = stored + 1
StepLeft = stored & ":" & Helper()
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Private stored As Long
Private Function Helper() As Long
Static count As Long
count = count + 1
Helper = count
End Function
Public Function StepRight() As String
stored = stored + 1
StepRight = stored & ":" & Helper()
End Function
