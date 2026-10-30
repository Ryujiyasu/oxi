Attribute VB_Name = "DriverModule"
Option Explicit
Public held As WorkerClass, Log As String
Public Property Get Selected() As WorkerClass
Log = Log & "R"
Set Selected = held
End Property
Public Function NextIndex() As Long
Log = Log & "I"
NextIndex = 1
End Function
Public Function NextValue() As Long
Log = Log & "V"
NextValue = 9
End Function
Public Function Probe() As String
Set held = New WorkerClass
Log = ""
Selected.Item(NextIndex()) = NextValue()
Probe = Log
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public Property Let Item(ByVal index As Long, ByVal value As Long)
DriverModule.Log = DriverModule.Log & "C"
End Property
