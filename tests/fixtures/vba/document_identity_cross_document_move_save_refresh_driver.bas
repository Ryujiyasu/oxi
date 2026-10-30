Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 If Ledger.Index = 1 Then Ledger.Move After:=Journal
 Probe = CStr(Ledger.Tick()) & ":" & CStr(Journal.Tick()) & ":" & CStr(ArchiveHost.Tick()) & ":" & CStr(Ledger.Index) & ":" & CStr(Journal.Index) & ":" & CStr(Ledger Is Worksheets(Ledger.Index))
End Function
