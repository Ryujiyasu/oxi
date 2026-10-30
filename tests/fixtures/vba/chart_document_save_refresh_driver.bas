Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 If PlotHost.Index = 1 Then PlotHost.Move After:=Ledger
 Probe = CStr(PlotHost.Tick()) & ":" & CStr(PlotHost.Index) & ":" & CStr(Worksheets(1).Index)
End Function
