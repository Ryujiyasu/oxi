Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Probe = "ran"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "BrokenClass"
Private Const broken As Byte = 300
