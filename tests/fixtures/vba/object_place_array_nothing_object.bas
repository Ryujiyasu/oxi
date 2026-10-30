Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim workers(0) As Object
On Error Resume Next
workers(0) = 7
Probe = CStr(Err.Number) & ":" & CStr(workers(0) Is Nothing)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public number As Long
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
number = value
End Property
Public Property Get Value() As Long
Value = number
End Property
