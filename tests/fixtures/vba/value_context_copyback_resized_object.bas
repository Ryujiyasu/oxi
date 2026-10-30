Attribute VB_Name = "DriverModule"
Option Explicit
Public workers() As WorkerClass
Public Log As String
Public Function Probe() As String
Log = ""
ReDim workers(0)
Set workers(0) = New WorkerClass
On Error Resume Next
Change workers(0)
Probe = CStr(Err.Number) & ":" & CStr(UBound(workers)) & ":" & Log & ":" & CStr(workers(0) Is Nothing)
End Function
Private Sub Change(ByRef value As Object)
value = 7
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
ReDim DriverModule.workers(0 To 1)
DriverModule.Log = DriverModule.Log & "S"
End Property
Private Sub Class_Terminate()
DriverModule.Log = DriverModule.Log & "T"
End Sub
