Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim text As String * 4
text = "abcd"
worker.Item(text) = 9
Probe = "[" & text & "]"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Let Item(ByRef index As String, ByVal value As Long)
index = "123456"
End Property
