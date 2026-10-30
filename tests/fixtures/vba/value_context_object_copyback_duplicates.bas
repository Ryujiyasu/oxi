Attribute VB_Name = "DriverModule"
Option Explicit
Public held As WorkerClass
Public seen As String
Public Function Probe() As String
Set held = New WorkerClass
On Error Resume Next
Change held, held
Probe = seen & ":" & CStr(held Is Nothing)
End Function
Private Sub Change(ByRef first As Object, ByRef second As Object)
Set first = Nothing
seen = CStr(second Is Nothing) & ":" & CStr(held Is Nothing)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
