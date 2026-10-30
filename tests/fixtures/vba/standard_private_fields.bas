Attribute VB_Name = "DriverModule"
Option Explicit
Public Function ReadBoth() As String
ReadBoth = CStr(StepLeft()) & "|" & CStr(StepRight())
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Private stored As Long
Public Function StepLeft() As Long
stored = stored + 1
StepLeft = stored
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Private stored As Long
Public Function StepRight() As Long
stored = stored + 1
StepRight = stored
End Function
