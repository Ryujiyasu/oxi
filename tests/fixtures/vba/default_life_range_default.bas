Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim target As Range
Set target = Range("A1")
target = 7
Probe = CStr(Range("A1").Value)
End Function
