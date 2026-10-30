Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim text As String * 4
Dim result As Variant
text = "abcd"
result = Len(text)
Probe = TypeName(result) & ":" & result
End Function
