Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function Probe() As String
Log = ""
Dim worker As New WorkerClass
Dim result As Long
result = worker(NextValue())
Probe = Log & ":" & CStr(result)
End Function
Public Function NextValue() As Long
Log = Log & "A"
NextValue = 1
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Private Sub Class_Initialize()
DriverModule.Log = DriverModule.Log & "I"
End Sub
Public Function Change(ByVal number As Long) As Long
Attribute Change.VB_UserMemId = 0
DriverModule.Log = DriverModule.Log & "C"
Change = number + 10
End Function
