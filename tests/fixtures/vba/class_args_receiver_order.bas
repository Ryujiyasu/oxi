Attribute VB_Name = "DriverModule"
Option Explicit
Public held As WorkerClass
Public Log As String
Public Property Get Selected() As WorkerClass
Log = Log & "R"
Set Selected = held
End Property
Public Function NextValue() As Long
Log = Log & "A"
NextValue = 1
End Function
Public Function Probe() As String
Set held = New WorkerClass
Log = ""
Selected.Change NextValue()
Probe = Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Sub Change(ByVal value As Long)
DriverModule.Log = DriverModule.Log & "C"
End Sub
