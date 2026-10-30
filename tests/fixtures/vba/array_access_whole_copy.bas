Attribute VB_Name = "DriverModule"
Option Explicit
Public items(0 To 0) As Long
Private Sub Change(ByRef values() As Long)
values(0) = values(0) + 10
End Sub
Public Function Probe() As String
Dim values(0 To 1) As Long, copied As Variant
values(0) = 3
copied = values()
values(0) = 9
Probe = CStr(copied(0)) & ":" & CStr(values(0))
End Function
