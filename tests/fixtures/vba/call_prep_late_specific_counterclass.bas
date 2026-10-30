Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public held As CounterClass
Public worker As WorkerClass
Public Function Probe() As String
Log = ""
Set worker = New WorkerClass
Set held = New CounterClass
Log = Log & "B"
Selected.Change held
Log = Log & "A"
Probe = Log & "|" & CStr(held Is Nothing)
End Function
Public Property Get Selected() As Object
Log = Log & "R"
Set held = Nothing
Log = Log & "G"
Set Selected = worker
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByRef value As CounterClass)
DriverModule.Log = DriverModule.Log & "C" & CStr(value Is Nothing)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Private Sub Class_Initialize()
DriverModule.Log = DriverModule.Log & "I"
End Sub
Private Sub Class_Terminate()
DriverModule.Log = DriverModule.Log & "T"
End Sub
