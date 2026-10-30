Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As Ledger
 Set s = New Ledger
 Probe = CStr(s.Helper(7))
End Function
