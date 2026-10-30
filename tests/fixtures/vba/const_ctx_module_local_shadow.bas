Attribute VB_Name = "DriverModule"
Option Explicit
Public Const size = 2
Public Function Probe() As String
On Error Resume Next
Dim size As Long
Dim value(size) As Long
Probe = "ran"
End Function
