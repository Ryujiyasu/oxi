Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
On Error Resume Next
Probe = Consume(Nothing)
End Function
Private Function Consume(ByVal value As Long) As String
Consume = TypeName(value)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
