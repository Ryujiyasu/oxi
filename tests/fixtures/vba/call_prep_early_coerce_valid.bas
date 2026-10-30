Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public worker As WorkerClass
Public Function Probe() As String
Dim value As Variant
value = "3"
Log = ""
Set worker = New WorkerClass
On Error Resume Next
Selected.Change value
Probe = CStr(Err.Number) & ":" & Log
End Function
Public Property Get Selected() As WorkerClass
Log = Log & "R"
Set Selected = worker
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Sub Change(ByVal value As Long)
DriverModule.Log = DriverModule.Log & "C" & TypeName(value) & CStr(value)
End Sub
