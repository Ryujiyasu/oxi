Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As WorkerClass
Set worker = New WorkerClass
Dim number As Long
number = 1
With worker
.Change number, number
End With
Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ByRef first As Long, ByRef second As Long)
first = first + 1
second = second + first
End Sub
