Attribute VB_Name = "DriverModule"
Option Explicit
Public calls As Long
Public Function NextIndex() As Long
calls = calls + 1
NextIndex = 0
End Function
Public Function Probe() As String
Dim numbers(0 To 0) As Long
Dim result As Variant
numbers(0) = 123
On Error Resume Next
result = Len(numbers(NextIndex()))
Probe = Err.Number & ":" & TypeName(result) & ":" & result & ":" & calls
End Function
