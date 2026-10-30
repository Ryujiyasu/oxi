Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim b As Workbook
 Set b = ThisWorkbook
 Dim n As Long
 n = 7
 b.Bump n
 Probe = CStr(n)
End Function
