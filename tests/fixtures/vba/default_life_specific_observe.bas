Attribute VB_Name = "DriverModule"
Option Explicit
Public held As CounterClass
Public seen As String
Public Function Probe() As String
Set held = New CounterClass
Change held
Probe = seen & ":" & CStr(held Is Nothing)
End Function
Private Sub Change(ByRef value As Variant)
Set value = Nothing
seen = CStr(held Is Nothing)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Public number As Long
