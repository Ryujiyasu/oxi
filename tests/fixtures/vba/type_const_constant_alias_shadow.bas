Attribute VB_Name = "DriverModule"
Option Explicit
Public Const size = 1
Public Const other = size
Public Function Probe() As String
On Error Resume Next
Const size = DriverModule.other
Dim value(size) As Long
Probe = CStr(UBound(value))
End Function
