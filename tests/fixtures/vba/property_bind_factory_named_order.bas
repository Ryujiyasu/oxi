Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public held As WorkerClass
Public Function Probe() As String
Set held = New WorkerClass
Log = ""
Selected().Change second:=NextValue("S"), first:=NextValue("F")
Probe = Log
End Function
Public Function Selected() As WorkerClass
Log = Log & "R"
Set Selected = held
End Function
Public Function NextValue(ByVal text As String) As Long
Log = Log & text
NextValue = 1
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByVal first As Long, ByVal second As Long)
DriverModule.Log = DriverModule.Log & "C"
End Sub
