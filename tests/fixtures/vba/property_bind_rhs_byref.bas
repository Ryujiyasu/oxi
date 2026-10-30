Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long, rhs As Long
number = 1
rhs = 3
worker.Item(number) = rhs
Probe = CStr(number) & ":[" & CStr(rhs) & "]"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Property Let Item(ByRef index As Long, ByRef value As Long)
index = index + 10
value = value + 10
End Property
