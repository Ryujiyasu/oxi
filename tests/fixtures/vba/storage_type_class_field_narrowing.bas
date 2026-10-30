Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As New WorkerClass
On Error Resume Next
worker.number = 2.6
Dim first As String
first = TypeName(worker.number) & ":" & CStr(worker.number)
Err.Clear
worker.number = "bad"
Probe = first & ":" & CStr(Err.Number) & ":" & CStr(worker.number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public number As Long
