Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long
number = 1
worker.Item(number) = 9
Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Property Let Item(ByRef index As Long, ByVal value As Long)
index = index + 10
End Property
