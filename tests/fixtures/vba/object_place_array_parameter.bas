Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim workers(0) As WorkerClass
Set workers(0) = New WorkerClass
On Error Resume Next
Change workers
Probe = CStr(Err.Number) & ":" & CStr(workers(0).number)
End Function
Private Sub Change(ByRef values() As WorkerClass)
values(0) = 7
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public number As Long
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
number = value
End Property
Public Property Get Value() As Long
Value = number
End Property
