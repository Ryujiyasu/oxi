Attribute VB_Name = "DriverModule"
Option Explicit
Public items As Variant
Public trace As String
Public Function Probe() As String
items = Array(1)
On Error Resume Next
Change items(0)
Probe = trace & ":" & items(0)
End Function
Private Sub Change(ByRef value As Variant)
On Error Resume Next
items = 3
trace = CStr(Err.Number)
value = value + 1
End Sub
