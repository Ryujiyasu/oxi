Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim n As Long
 n = 7
 Ledger.Bump n
 Probe = CStr(n)
End Function
