Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim values(0 To 0) As Long
values(0) = 1
Change values()
Probe = CStr(values(0))
End Function
Private Sub Change(ByRef values() As Long)
values(0) = values(0) + 10
End Sub
