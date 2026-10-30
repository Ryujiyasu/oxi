Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function Probe() As String
Dim worker As WorkerClass
Set worker = New WorkerClass
Log = ""
On Error Resume Next
worker.Hidden
Probe = CStr(Err.Number) & ":" & Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Friend Sub Hidden()
DriverModule.Log = "ok"
End Sub
