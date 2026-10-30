Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function Probe() As String
Dim worker As New WorkerClass, counter As Long
counter = 1
Log = ""
worker.SelectItem(counter).Touch
Probe = CStr(counter) & ":" & Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Function SelectItem(ByRef value As Long) As WorkerClass
value = value + 1
Set SelectItem = Me
End Function
Public Sub Touch()
DriverModule.Log = DriverModule.Log & "X"
End Sub
Public Sub Change(ByRef value As Long)
value = value + 10
DriverModule.Log = DriverModule.Log & "X"
End Sub
