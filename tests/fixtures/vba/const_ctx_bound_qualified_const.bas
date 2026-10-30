Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(SizesModule.Low) As Long
Probe = CStr(Err.Number) & ":ran"
End Function
Attribute VB_Name = "SizesModule"
Public Const Low = -1
