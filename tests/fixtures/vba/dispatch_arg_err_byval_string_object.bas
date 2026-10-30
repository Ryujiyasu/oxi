Attribute VB_Name = "DriverModule"
Option Explicit
Public worker As WorkerClass
Public Log As String
Public Function Probe() As String
Set worker = New WorkerClass
Err.Number = 5
Selected.Change Err
Probe = Log
End Function
Public Property Get Selected() As Object
Err.Number = 7
Set Selected = worker
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByVal value As String)
DriverModule.Log = CStr(value) & ":" & CStr(Err.Number)
End Sub
