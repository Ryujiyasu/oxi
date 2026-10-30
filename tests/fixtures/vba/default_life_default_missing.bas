Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As WorkerClass
Set worker = New WorkerClass
On Error Resume Next
worker = 7
Probe = CStr(Err.Number) & ":" & CStr(worker.number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
