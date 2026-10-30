Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim obj As New Sample
Probe = CStr(obj.NextValue())
End Function
Attribute VB_Name = "CounterModule"
Public count As Long
VERSION 1.0 CLASS
Attribute VB_Name = "Sample"
Public Function NextValue() As Long
count = count + 1
NextValue = count
End Function
