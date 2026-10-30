Attribute VB_Name = "DriverModule"
Option Explicit
Public Type Sample
item As Long
End Type
Public Function Probe() As String
On Error Resume Next
Const value As Sample = 1
Probe = "ran"
End Function
