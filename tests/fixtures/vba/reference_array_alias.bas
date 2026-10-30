Attribute VB_Name = "DriverModule"
Option Explicit
Public trace As String
Public Function Probe() As String
LeftModule.items(1) = 1
RightModule.Change LeftModule.items(1), LeftModule.items(1)
Probe = trace & "|" & LeftModule.items(1)
End Function
Attribute VB_Name = "LeftModule"
Option Explicit
Public items(1 To 1) As Long
Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Change(ByRef first As Long, ByRef second As Long)
first = first + 1
trace = first & ":" & second & ":" & LeftModule.items(1)
second = second + 10
End Sub
