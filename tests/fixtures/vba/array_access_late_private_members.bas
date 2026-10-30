Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As Object
Set worker = New WorkerClass
On Error Resume Next
worker.number = 7
Dim first As String, result As String
first = CStr(Err.Number) & ":" & CStr(worker.ReadNumber)
Err.Clear
result = worker.Hidden()
Probe = first & ":" & CStr(Err.Number) & ":" & result
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Private number As Long
Public Property Get ReadNumber() As Long
ReadNumber = number
End Property
Private Function Hidden() As String
Hidden = "ok"
End Function
