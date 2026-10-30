Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As WorkerClass
Set worker = New WorkerClass
Dim number As Long
number = 1
worker.Item((number), number) = 9
Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Let Item(ByRef first As Long, ByRef second As Long, ByVal value As Long)
first = first + 1
second = second + first
End Property
