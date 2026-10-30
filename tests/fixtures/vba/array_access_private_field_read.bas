Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As Object
Set worker = New WorkerClass
On Error Resume Next
Dim result As Long
result = worker.number
Probe = CStr(Err.Number) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Private number As Long
