Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim result As Variant
result = Len(worker.Number())
Probe = TypeName(result) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Function Number() As Long
Number = 123
End Function
