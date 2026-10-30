Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim values() As Long, original(0 To 0) As Long
original(0) = 3
values = original
On Error Resume Next
ReDim Preserve values(0 To 1)
Probe = CStr(Err.Number) & ":" & CStr(UBound(values)) & ":" & CStr(values(0))
End Function
