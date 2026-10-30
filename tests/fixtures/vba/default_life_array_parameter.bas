Attribute VB_Name = "DriverModule"
Option Explicit
Private Sub CopyArray(ByRef target() As Long, ByRef source() As Long)
target = source
End Sub
Public Function Probe() As String
Dim values() As Long, original(0 To 0) As Long
ReDim values(0 To 0)
original(0) = 3
CopyArray values, original
On Error Resume Next
ReDim Preserve values(0 To 1)
Probe = CStr(Err.Number) & ":" & CStr(UBound(values)) & ":" & CStr(values(0))
End Function
