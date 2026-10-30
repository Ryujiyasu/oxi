Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim leftValue As String * 3
Dim rightValue As String * 3
rightValue = "ab"
LSet leftValue = rightValue
Probe = leftValue
End Function
