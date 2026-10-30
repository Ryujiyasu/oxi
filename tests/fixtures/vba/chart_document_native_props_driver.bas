Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(PlotHost.ChartType) & ":" & CStr(PlotHost.HasTitle) & ":" & CStr(PlotHost.HasLegend) & ":" & CStr(PlotHost.SeriesCollection.Count)
End Function
