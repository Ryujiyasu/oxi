Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As Ledger
 Set s = Ledger
 Probe = CStr(s.Helper(7)) & ":" & CStr(s Is Ledger)
End Function
