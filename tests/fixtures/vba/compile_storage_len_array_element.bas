Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim numbers(0 To 0) As Long
numbers(0) = 123
On Error Resume Next
Dim result As Variant
result = Len(numbers(0))
Probe = Err.Number & ":" & TypeName(result)
If Not IsNull(result) Then Probe = Probe & ":" & CStr(result)
End Function
