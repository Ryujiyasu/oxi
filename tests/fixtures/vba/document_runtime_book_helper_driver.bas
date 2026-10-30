Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(ThisWorkbook.Helper(7))
End Function
