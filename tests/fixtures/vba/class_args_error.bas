Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long
number = 1
On Error Resume Next
worker.Change number
Probe = CStr(Err.Number) & ":" & CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ByRef number As Long)
number = 9
Err.Raise 513
End Sub
