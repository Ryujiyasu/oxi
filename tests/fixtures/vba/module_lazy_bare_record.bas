Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
point.value = point.value + 1
Probe = CStr(point.value)
End Function
Attribute VB_Name = "CounterModule"
Public Type PointType
value As Long
End Type
Public point As PointType
