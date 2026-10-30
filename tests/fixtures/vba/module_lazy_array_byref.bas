Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Bump values
Probe = CStr(values(1))
End Function
Private Sub Bump(ByRef arg() As Long)
arg(1) = arg(1) + 2
End Sub
Attribute VB_Name = "CounterModule"
Public values(1 To 2) As Long
