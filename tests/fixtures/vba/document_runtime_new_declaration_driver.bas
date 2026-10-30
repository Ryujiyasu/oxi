Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As New Ledger
 Probe = CStr(s.Helper(7))
End Function
