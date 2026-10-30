Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public Function Probe() As String
number = 1
Probe = CStr(number = "1") & ":" & Compare(number)
End Function
Private Function Compare(ByRef value As Variant) As String
Compare = CStr(value = "1") & ":" & CStr(value < "2")
End Function
