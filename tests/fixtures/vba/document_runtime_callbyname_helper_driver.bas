Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(CallByName(Ledger, "Helper", VbMethod, 7))
End Function
