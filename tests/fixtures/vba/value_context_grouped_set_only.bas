Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim worker As New WorkerClass, value As Variant
On Error Resume Next
value = (worker)
Probe = CStr(Err.Number) & ":" & TypeName(value)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Property Set Value(ByVal value As Object)
Attribute Value.VB_UserMemId = 0
End Property
