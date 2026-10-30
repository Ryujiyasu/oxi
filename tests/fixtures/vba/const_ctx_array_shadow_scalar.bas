Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim Len As Long
Dim value(Len("a")) As Long
Probe = "ran"
End Function
