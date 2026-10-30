Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As Worksheet
 Set s = Ledger.NativeReference()
 Probe = TypeName(s) & ":" & CStr(s Is Ledger)
End Function
