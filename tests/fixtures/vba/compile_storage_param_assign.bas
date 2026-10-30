Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = Exercise(Array(1, 2))
End Function
Private Function Exercise(ParamArray values() As Variant) As String
Dim copied As Variant
On Error Resume Next
values = Array(3, 4)
Exercise = CStr(Err.Number)
End Function
