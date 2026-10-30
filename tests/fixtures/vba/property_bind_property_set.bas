Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass, other As New WorkerClass
Dim number As Long
number = 1
Set worker.Item(number) = other
Probe = CStr(number) & ":" & CStr(worker.held Is other)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public held As WorkerClass
Public Property Set Item(ByRef index As Long, ByVal value As WorkerClass)
index = index + 10
Set held = value
End Property
