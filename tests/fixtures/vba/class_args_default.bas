Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long, result As Long
number = 1
result = worker(number)
Probe = CStr(number) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Function Change(ByRef number As Long) As Long
Attribute Change.VB_UserMemId = 0
number = number + 10
Change = number
End Function
