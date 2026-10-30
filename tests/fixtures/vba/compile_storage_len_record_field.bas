Attribute VB_Name = "DriverModule"
Option Explicit
Public Type Entry
number As Long
End Type
Public Function Probe() As String
Dim held As Entry
held.number = 123
On Error Resume Next
Dim result As Variant
result = Len(held.number)
Probe = Err.Number & ":" & TypeName(result)
If Not IsNull(result) Then Probe = Probe & ":" & CStr(result)
End Function
