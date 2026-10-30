Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public held As CounterClass
Public Function Probe() As String
Log = ""
Set held = New CounterClass
Release held
Probe = Log & "|" & CStr(held Is Nothing)
End Function
Private Sub Release(ParamArray values() As Variant)
Log = Log & "B"
Set values(0) = Nothing
values = Array(1)
Log = Log & "A"
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Private Sub Class_Initialize()
Log = Log & "I"
End Sub
Private Sub Class_Terminate()
Log = Log & "T"
End Sub
