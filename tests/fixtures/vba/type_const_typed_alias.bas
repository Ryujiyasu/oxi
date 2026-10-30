Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const seed = 300
Const value As Byte = seed
Probe = "ran"
End Function
