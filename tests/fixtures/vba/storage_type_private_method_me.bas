Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Probe = worker.Length
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Function Length() As String
Length = Me.Hidden()
End Function
Private Function Hidden() As String
Hidden = "ok"
End Function
