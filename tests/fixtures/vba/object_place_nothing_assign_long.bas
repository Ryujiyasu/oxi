Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim value As Long
On Error Resume Next
value = Nothing
Probe = CStr(Err.Number)
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
