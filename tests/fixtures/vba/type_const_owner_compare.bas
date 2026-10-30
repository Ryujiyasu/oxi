Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(SizesModule.Flag + 1) As Long
Probe = CStr(UBound(value))
End Function
Attribute VB_Name = "SizesModule"
Option Compare Text
Public Const Flag = ("a" = "A")
