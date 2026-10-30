Attribute VB_Name = "DriverModule"
Option Explicit
Public items() As Long
Public trace As String
Public Function Probe() As String
ReDim items(1 To 1)
items(1) = 1
On Error Resume Next
Change items(1)
Err.Clear
ReDim items(1 To 2)
Probe = Err.Number & ":" & UBound(items)
End Function
Private Sub Change(ByRef value As Long)
value = value + 1
Err.Raise 513
trace = CStr(Err.Number)
value = value + 1
End Sub
