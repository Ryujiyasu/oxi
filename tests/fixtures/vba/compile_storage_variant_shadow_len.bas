Attribute VB_Name = "DriverModule"
Option Explicit
Public value As Long
Public Function Probe() As String
value = 123
Probe = Inspect(value)
End Function
Private Function Inspect(ByRef value As Variant) As String
Dim result As Variant
result = Len(value)
Inspect = TypeName(result) & ":" & result
End Function
