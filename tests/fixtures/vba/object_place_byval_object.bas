Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass
On Error Resume Next
Change worker
Probe = CStr(Err.Number) & ":" & CStr(worker.number)
End Function
Private Sub Change(ByVal value As Object)
value = 7
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
