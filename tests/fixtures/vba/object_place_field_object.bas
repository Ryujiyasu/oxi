Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim box As New HolderClass
Set box.worker = New WorkerClass
On Error Resume Next
box.worker = 7
Probe = CStr(Err.Number) & ":" & TypeName(box.worker)
If IsObject(box.worker) Then Probe = Probe & ":" & CStr(box.worker.number)
End Function
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
VERSION 1.0 CLASS
Attribute VB_Name = "HolderClass"
Public worker As Object
