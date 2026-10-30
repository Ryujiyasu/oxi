Attribute VB_Name = "DriverModule"
Option Explicit
Public held As Collection
Public worker As WorkerClass
Public Log As String
Public Function Probe() As String
Log = ""
Set held = New Collection
held.Add New CounterClass
Set worker = New WorkerClass
On Error Resume Next
Selected.Change held
Probe = CStr(Err.Number) & ":" & Log & "|" & CStr(held Is Nothing)
End Function
Public Property Get Selected() As Object
Log = Log & "R"
Set held = Nothing
Log = Log & "G"
Set Selected = worker
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByVal value As Object)
DriverModule.Log = DriverModule.Log & "C" & CStr(value Is Nothing)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Private Sub Class_Terminate()
DriverModule.Log = DriverModule.Log & "T"
End Sub
