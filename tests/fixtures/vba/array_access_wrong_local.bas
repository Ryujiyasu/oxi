Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim target As FirstClass, source As Object
Set source = New OtherClass
On Error Resume Next
Set target = source
Probe = CStr(Err.Number) & ":" & CStr(target Is Nothing)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
VERSION 1.0 CLASS
Attribute VB_Name = "FirstClass"
Implements IFirst
Private Function IFirst_Read() As Long
IFirst_Read = 1
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "OtherClass"
Public number As Long
VERSION 1.0 CLASS
Attribute VB_Name = "IFirst"
Public Function Read() As Long
End Function
