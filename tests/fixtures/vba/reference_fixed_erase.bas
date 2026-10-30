Attribute VB_Name = "DriverModule"
Option Explicit
Public items(1 To 1) As Long
Public trace As String
Public Function Probe() As String
items(1) = 1
On Error Resume Next
Change items(1)
Probe = trace & ":" & items(1)
End Function
Private Sub Change(ByRef value As Long)
On Error Resume Next
Erase items
trace = CStr(Err.Number)
value = value + 1
End Sub
