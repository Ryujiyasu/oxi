Attribute VB_Name = "DriverModule"
Option Explicit
Public number As LongLong
Public Function Probe() As String
number = 1
Probe = Change(number) & "|" & number
End Function
Private Function Change(ParamArray values() As Variant) As String
On Error Resume Next
values(0) = 2.6
Change = Err.Number & ":" & number
End Function
