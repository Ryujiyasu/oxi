Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const seed As Byte = 300
Const value = seed
Probe = "ran"
End Function
