Attribute VB_Name = "DriverModule"
Option Explicit
Public values As Variant
Public trace As String
Public calls As Long
Public Function Index() As Long
calls = calls + 1
Index = 1
End Function
Public Function Probe() As String
values = Array(Array(1, 2))
Change values(0)(Index()), values(0)(Index())
Probe = trace & "|" & values(0)(1) & ":" & calls
End Function
Private Sub Change(ByRef first As Variant, ByRef second As Variant)
first = first + 1
trace = first & ":" & second & ":" & values(0)(1)
second = second + 10
End Sub
