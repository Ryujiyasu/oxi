Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Bump count
Probe = CStr(count)
End Function
Private Sub Bump(ByRef value As Long)
value = value + 3
End Sub
Attribute VB_Name = "CounterModule"
Public count As Long
