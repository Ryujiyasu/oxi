Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim first As New WorkerClass, second As New WorkerClass
first.number = 1
first.Forward second
Probe = CStr(first.number) & ":" & CStr(second.number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
Public Sub Forward(ByVal other As WorkerClass)
other.Change number
End Sub
Public Sub Change(ByRef value As Long)
number = 50
value = value + 10
End Sub
