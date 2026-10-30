Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public counters() As CounterClass
Public worker As WorkerClass
Public Function Probe() As String
Log = ""
ReDim counters(0)
Set counters(0) = New CounterClass
Set worker = New WorkerClass
On Error Resume Next
Selected.Change counters(0)
Probe = CStr(Err.Number) & ":" & CStr(UBound(counters)) & ":" & Log
End Function
Public Property Get Selected() As Object
Log = Log & "R"
ReDim counters(0 To 1)
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
