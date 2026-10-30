Attribute VB_Name = "DriverModule"
Option Explicit
Public held As CounterClass
Public Function Probe() As String
Set held = New CounterClass
On Error Resume Next
Change held
Probe = CStr(Err.Number) & ":" & CStr(held Is Nothing)
End Function
Private Sub Change(ByRef value As Variant)
Set value = Nothing
Err.Raise 513
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Public number As Long
