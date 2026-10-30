Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim text As String * 4
text = "abcd"
worker.Change text
Probe = "[" & text & "]"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ByRef text As String)
text = "123456"
End Sub
