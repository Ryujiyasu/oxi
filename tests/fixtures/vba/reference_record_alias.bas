Attribute VB_Name = "DriverModule"
Option Explicit
Public trace As String
Public Function Probe() As String
LeftModule.held.amount = 1
RightModule.Change LeftModule.held.amount, LeftModule.held.amount
Probe = trace & "|" & LeftModule.held.amount
End Function
Attribute VB_Name = "LeftModule"
Option Explicit
Public Type Entry
amount As Long
End Type
Public held As Entry
Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Change(ByRef first As Long, ByRef second As Long)
first = first + 1
trace = first & ":" & second & ":" & LeftModule.held.amount
second = second + 10
End Sub
