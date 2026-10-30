Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(Ledger.Tick()) & ":" & CStr(Journal.Tick()) & ":" & CStr(ArchiveHost.Tick()) & ":" & CStr(Ledger.Index) & ":" & CStr(Journal.Index) & ":" & CStr(Ledger Is Worksheets(Ledger.Index))
End Function
