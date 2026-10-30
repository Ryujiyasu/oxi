Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 If Ledger Is Worksheets(1) Then Ledger.Move After:=Worksheets(2)
 Probe = Ledger.Name & ":" & Ledger.CodeName & ":" & CStr(Ledger.Tick()) & ":" & CStr(Ledger Is Worksheets(2))
End Function
