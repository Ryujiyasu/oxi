Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As WorkerClass
Set worker = New WorkerClass
Dim result As String
On Error Resume Next
result = worker.Hidden
Probe = CStr(Err.Number) & ":" & result
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Friend Property Get Hidden() As String
Hidden = "ok"
End Property
