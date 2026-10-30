Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function Probe() As String
Log = ""
Dim held As CounterClass
Set held = New CounterClass
Log = Log & "B"
Change held
Log = Log & "A"
Probe = Log & "|" & CStr(held Is Nothing)
End Function
Private Sub Change(ByRef value As Variant)
Set value = Nothing
Log = Log & "C"
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Private Sub Class_Initialize()
DriverModule.Log = DriverModule.Log & "I"
End Sub
Private Sub Class_Terminate()
DriverModule.Log = DriverModule.Log & "T"
End Sub
