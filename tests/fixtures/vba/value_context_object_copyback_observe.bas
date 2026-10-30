Attribute VB_Name = "DriverModule"
Option Explicit
Public held As WorkerClass
Public seen As String
Public Function Probe() As String
Set held = New WorkerClass
On Error Resume Next
Change held
Probe = seen & ":" & CStr(held Is Nothing)
End Function
Private Sub Change(ByRef value As Object)
Set value = Nothing
seen = CStr(held Is Nothing)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
