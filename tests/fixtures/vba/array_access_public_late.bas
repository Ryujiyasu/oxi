Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As Object
Set worker = New WorkerClass
On Error Resume Next
Dim result As String
result = worker.Hidden()
Probe = CStr(Err.Number) & ":" & result
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Function Hidden() As String
Hidden = "ok"
End Function
