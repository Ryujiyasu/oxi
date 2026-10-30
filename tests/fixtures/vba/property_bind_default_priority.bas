Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long
number = 3
Probe = CStr(worker(number)) & ":" & CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Function DefaultValue(ByRef number As Long) As Long
Attribute DefaultValue.VB_UserMemId = 0
number = number + 10
DefaultValue = number
End Function
Public Function Item(ByVal number As Long) As Long
Item = number + 100
End Function
