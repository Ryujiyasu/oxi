Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function NextValue(ByVal text As String) As Long
Log = Log & text
NextValue = 1
End Function
Public Function Probe() As String
Dim worker As WorkerClass
Set worker = New WorkerClass
Log = ""
worker.Item(second:=NextValue("S"),first:=NextValue("F")) = NextValue("V")
Probe = Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Let Item(ByVal first As Long, ByVal second As Long, ByVal value As Long)
DriverModule.Log = DriverModule.Log & "C"
End Property
