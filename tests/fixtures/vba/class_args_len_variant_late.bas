Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As Object
Set worker = New WorkerClass
worker.value = 123
Dim result As Variant
result = Len(worker.value)
Probe = TypeName(result) & ":" & CStr(result)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public value As Variant
