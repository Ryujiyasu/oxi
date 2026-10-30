Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As IWorker
Set worker = New WorkerClass
Probe = worker.Change()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Implements IWorker
Public Function Change() As String
Change = "own"
End Function
Private Function IWorker_Change() As String
IWorker_Change = "interface"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "IWorker"
Public Function Change() As String
End Function
