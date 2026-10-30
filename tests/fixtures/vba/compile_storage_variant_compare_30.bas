Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim number As Long
number = 30
Probe = CStr(number > "9") & ":" & Compare(number)
End Function
Private Function Compare(ByRef value As Variant) As String
Compare = CStr(value > "9")
End Function
