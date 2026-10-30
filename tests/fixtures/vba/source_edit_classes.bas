Option Explicit
Public Log As String
Private held As StateClass
Private other As Object
Public Sub Prime()
Set held = New StateClass
Set other = held
held.Value = 123
Range("D1").Value = ReadState()
End Sub
Public Function ReadState() As String
If held Is Nothing Then
ReadState = "none|" & Log
Else
ReadState = held.ReadValue() & "|" & CStr(held Is other) & "|" & Log
End If
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "StateClass"
Option Explicit
Public Value As Long
Private Sub Class_Initialize()
Log = Log & "I"
End Sub
Private Sub Class_Terminate()
Log = Log & "T"
End Sub
Public Function ReadValue() As String
Static stored As Long
stored = stored + 1
ReadValue = CStr(Value) & "|" & CStr(stored)
End Function
