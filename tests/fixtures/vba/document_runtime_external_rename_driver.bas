Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = Ledger.Name & ":" & CStr(Ledger.Tick())
End Function
