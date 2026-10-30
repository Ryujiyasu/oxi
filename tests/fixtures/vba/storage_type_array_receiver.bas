Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function NextValue(ByVal text As String) As Long
Log = Log & text
NextValue = 1
End Function
Public Function Probe() As String
Dim workers(0 To 0) As WorkerClass
Set workers(0) = New WorkerClass
Log = ""
workers(0).Change second:=NextValue("S"), first:=NextValue("F")
Probe = Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ByVal first As Long, ByVal second As Long)
DriverModule.Log = DriverModule.Log & "C"
End Sub
