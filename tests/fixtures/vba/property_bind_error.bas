Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long
number = 1
On Error Resume Next
worker.Item(number) = 9
Probe = CStr(Err.Number) & ":" & CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Let Item(ByRef index As Long, ByVal value As Long)
index = 11
Err.Raise 513
End Property
