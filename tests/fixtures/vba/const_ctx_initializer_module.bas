Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value = SizesModule.High
Probe = CStr(value)
End Function
Attribute VB_Name = "SizesModule"
Public Const High = 2
