Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Probe = "ran"
End Function
Attribute VB_Name = "BadModule"
Private broken(-1) As Long
