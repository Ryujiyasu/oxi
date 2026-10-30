Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Change OtherModule.items()
Probe = CStr(OtherModule.items(0))
End Function
Private Sub Change(ByRef values() As Long)
values(0) = values(0) + 10
End Sub
Attribute VB_Name = "OtherModule"
Option Explicit
Public items(0 To 0) As Long
