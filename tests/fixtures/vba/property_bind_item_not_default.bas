Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
On Error Resume Next
Dim result As Long
result = worker(3)
Probe = CStr(Err.Number) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Function Item(ByVal number As Long) As Long
Item = number + 100
End Function
