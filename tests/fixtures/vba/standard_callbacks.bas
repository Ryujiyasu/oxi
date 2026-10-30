Attribute VB_Name = "DriverModule"
Option Explicit
Private worker As CounterClass
Public Function Probe() As String
If worker Is Nothing Then Set worker = New CounterClass
Probe = LeftModule.NextCount() & "|" & worker.StepOnce() & "|" & RightModule.NextCount()
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public Function NextCount() As Long
Static count As Long
count = count + 1
NextCount = count
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Public Function NextCount() As Long
Static count As Long
count = count + 1
NextCount = count
End Function

VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Option Explicit
Public Function StepOnce() As String
Static count As Long
count = count + 1
StepOnce = count & ":" & LeftModule.NextCount() & ":" & RightModule.NextCount()
End Function
