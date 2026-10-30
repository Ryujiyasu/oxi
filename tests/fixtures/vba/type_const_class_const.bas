Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As Sample = 1
Probe = "ran"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Sample"
Public item As Long
