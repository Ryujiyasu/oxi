Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As IWorker
Set worker = New WorkerClass
Dim number As Long
number = 1
worker.Change number
Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Implements IWorker
Private Sub IWorker_Change(ByRef number As Long)
number = number + 10
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "IWorker"
Public Sub Change(ByRef number As Long)
End Sub
