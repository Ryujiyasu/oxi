Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = PlotHost.Name & ":" & CStr(PlotHost.Tick()) & ":" & CStr(PlotHost.Index) & ":" & CStr(Worksheets(1).Index)
End Function
