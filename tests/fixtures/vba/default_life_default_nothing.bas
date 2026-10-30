Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As WorkerClass

On Error Resume Next
worker = 7
Probe = CStr(Err.Number) & ":" & CStr(worker Is Nothing)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
number = value
End Property
