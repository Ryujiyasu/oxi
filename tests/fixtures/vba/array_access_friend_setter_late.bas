Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As Object
Set worker = New WorkerClass
On Error Resume Next
worker.Value = 7
Probe = CStr(Err.Number) & ":" & CStr(worker.ReadNumber)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Private number As Long
Friend Property Let Value(ByVal value As Long)
number = value
End Property
Public Property Get ReadNumber() As Long
ReadNumber = number
End Property
