Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim numbers(0 To 1) As Long
numbers(0) = 1
worker.Change numbers
Probe = CStr(numbers(0))
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ByRef numbers() As Long)
numbers(0) = 11
End Sub
