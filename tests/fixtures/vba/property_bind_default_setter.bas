Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
Dim number As Long
number = 1
worker(number) = 9
Probe = CStr(number)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Let DefaultValue(ByRef index As Long, ByVal value As Long)
Attribute DefaultValue.VB_UserMemId = 0
index = index + 10
End Property
