Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As New WorkerClass
Probe = worker.Length
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Private numbers(0 To 1) As Long
Public Function Length() As String
numbers(0) = 123
Dim result As Variant
result = Len(numbers(0))
Length = TypeName(result) & ":" & CStr(result)
End Function
