Attribute VB_Name = "DriverModule"
Option Explicit
Public items(0 To 0) As Long
Private Sub Change(ByRef values() As Long)
values(0) = values(0) + 10
End Sub
Public Function Probe() As String
Dim values() As Long
Probe = CStr(IsArray(values()))
End Function
