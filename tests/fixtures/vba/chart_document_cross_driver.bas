Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(PlotHost.Tick()) & ":" & CStr(Ledger.Tick()) & ":" & CStr(ArchiveHost.Tick())
End Function
