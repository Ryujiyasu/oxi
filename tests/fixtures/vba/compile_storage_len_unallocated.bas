Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim numbers() As Long
Dim result As Variant
On Error Resume Next
result = Len(numbers(999))
Probe = Err.Number & ":" & TypeName(result) & ":" & result
End Function
