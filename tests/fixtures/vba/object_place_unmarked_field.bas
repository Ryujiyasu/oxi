Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As New WorkerClass
On Error Resume Next
worker = 7
Probe = CStr(Err.Number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Value As Long
