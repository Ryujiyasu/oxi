Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Set Ledger.Held = Ledger.Range("A1")
 Ledger.Held.Value = 7
 Probe = TypeName(Ledger.Held) & ":" & CStr(Ledger.Held.Value)
End Function
