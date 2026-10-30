Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public Function Probe() As String
number = 1
Probe = Change(number) & "|" & number
End Function
Private Function Change(ByVal value As Variant) As String
value = 2.6
Change = TypeName(value)
value = "bad"
Change = Change & ":" & TypeName(value)
End Function
