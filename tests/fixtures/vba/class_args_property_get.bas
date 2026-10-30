Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long, result As Long
number = 1
result = worker.Change(number)
Probe = CStr(number) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Get Change(ByRef number As Long) As Long
number = number + 10
Change = number
End Property
