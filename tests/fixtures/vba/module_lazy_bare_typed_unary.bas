Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
count = -32768
Dim answer As Variant
answer = -count
Probe = CStr(Err.Number)
End Function
Attribute VB_Name = "CounterModule"
Public count As Integer
