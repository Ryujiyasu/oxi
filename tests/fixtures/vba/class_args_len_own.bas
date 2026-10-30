Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Probe = worker.Length
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
Public Function Length() As String
number = 123
Dim result As Variant
result = Len(number)
Length = TypeName(result) & ":" & CStr(result)
End Function
