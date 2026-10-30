Attribute VB_Name = "DriverModule"
Option Explicit
Private Type Buffer
values(0 To 0) As Long
End Type
Public Function Probe() As String
Dim record As Buffer
record.values(0) = 1
Change record.values()
Probe = CStr(record.values(0))
End Function
Private Sub Change(ByRef values() As Long)
values(0) = values(0) + 10
End Sub
