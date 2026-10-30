Attribute VB_Name = "DriverModule"
Option Explicit
Private Const size = -1
Public Function Probe() As String
On Error Resume Next
Dim value(size) As Long
Probe = "ran"
End Function
