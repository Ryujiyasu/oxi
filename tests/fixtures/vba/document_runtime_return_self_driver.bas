Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As Ledger
 Set s = Ledger.SelfReference()
 Probe = TypeName(s) & ":" & CStr(s Is Ledger)
End Function
