Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long
number = 1
worker.Change number, number
Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ParamArray values() As Variant)
values(0) = 2
values(1) = values(1) + 10
End Sub
