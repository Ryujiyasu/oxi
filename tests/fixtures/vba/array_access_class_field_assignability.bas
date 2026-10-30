Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim worker As New WorkerClass
Dim other As Object
Set other = New OtherClass
On Error Resume Next
Set worker.held = other
Probe = CStr(Err.Number) & ":" & CStr(worker.held Is Nothing)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public held As FirstClass
VERSION 1.0 CLASS
Attribute VB_Name = "FirstClass"
Public number As Long
VERSION 1.0 CLASS
Attribute VB_Name = "OtherClass"
Public number As Long
