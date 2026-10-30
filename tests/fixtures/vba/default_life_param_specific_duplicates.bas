Attribute VB_Name = "DriverModule"
Option Explicit
Public held As CounterClass
Public seen As String
Public Function Probe() As String
Set held = New CounterClass
Change held, held
Probe = seen & ":" & CStr(held Is Nothing)
End Function
Private Sub Change(ParamArray values() As Variant)
Set values(0) = Nothing
seen = CStr(values(1) Is Nothing) & ":" & CStr(held Is Nothing)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Public number As Long
