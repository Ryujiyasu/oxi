Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim workers(0) As Object, boxed As Variant
Set workers(0) = New WorkerClass
boxed = workers
On Error Resume Next
boxed(0) = 7
Probe = CStr(Err.Number) & ":" & TypeName(boxed(0)) & ":" & CStr(workers(0).number)
If IsObject(boxed(0)) Then Probe = Probe & ":" & CStr(boxed(0).number)
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
