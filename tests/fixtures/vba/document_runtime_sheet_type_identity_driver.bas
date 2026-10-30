Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = TypeName(Ledger) & ":" & CStr(Ledger Is Worksheets(1)) & ":" & CStr(ObjPtr(Ledger) = ObjPtr(Worksheets(1)))
End Function
