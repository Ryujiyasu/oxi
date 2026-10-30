Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public Function Probe() As String
Dim worker As New WorkerClass
number = 1
worker.Change number
Probe = CStr(number) & ":" & CStr(worker.number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
Public Sub Change(ByRef value As Long)
number = 50
value = value + 10
End Sub
