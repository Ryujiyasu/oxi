Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
worker.Item() = 9
Probe = worker.Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Log As String
Public Property Let Item(Optional ByVal index As Long = 3, ByVal value As Long)
Log = CStr(index) & ":" & CStr(value)
End Property
