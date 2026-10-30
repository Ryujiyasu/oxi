Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 On Error Resume Next
 Dim n As Long
 n = Application.ThisWorkbook.Helper(7)
 Probe = CStr(Err.Number) & ":" & CStr(n)
End Function
