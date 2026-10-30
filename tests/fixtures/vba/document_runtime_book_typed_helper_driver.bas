Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim b As Workbook
 Set b = ThisWorkbook
 On Error Resume Next
 Dim n As Long
 n = b.Helper(7)
 Probe = CStr(Err.Number) & ":" & CStr(n)
End Function
