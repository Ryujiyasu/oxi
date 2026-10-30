Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Ledger.Counter = Ledger.Counter + 1
 Probe = CStr(Ledger.Counter)
End Function
