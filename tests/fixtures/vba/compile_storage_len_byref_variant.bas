Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim number As Long
number = 123
Probe = SizeOf(number)
End Function
Private Function SizeOf(ByRef value As Variant) As String
Dim result As Variant
result = Len(value)
SizeOf = TypeName(result) & ":" & result
End Function
