Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim result As Variant
worker.numbers(0) = 123
result = Len(worker.numbers(0))
Probe = TypeName(result) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public numbers(0 To 1) As Long
