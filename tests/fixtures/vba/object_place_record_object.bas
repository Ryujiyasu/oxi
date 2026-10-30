Attribute VB_Name = "DriverModule"
Option Explicit
Private Type Holder
worker As Object
End Type
Public Function Probe() As String
Dim box As Holder
Set box.worker = New WorkerClass
On Error Resume Next
box.worker = 7
Probe = CStr(Err.Number) & ":" & CStr(box.worker.number)
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
