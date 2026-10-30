Attribute VB_Name = "DriverModule"
Option Explicit
Public held As Range
Public worker As WorkerClass
Public Log As String
Public Function Probe() As String
Log = ""
Range("A1").Value = 7
Set held = Range("A1")
Set worker = New WorkerClass
On Error Resume Next
Selected.Change held
Probe = CStr(Err.Number) & ":" & Log & "|" & CStr(held Is Nothing)
End Function
Public Property Get Selected() As WorkerClass
Log = Log & "R"
Set held = Nothing
Log = Log & "G"
Set Selected = worker
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByRef value As Object)
DriverModule.Log = DriverModule.Log & "C" & CStr(value Is Nothing)
End Sub
