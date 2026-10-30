Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 PlotHost.HasTitle = True
 PlotHost.ChartTitle.Text = "Revenue"
 PlotHost.HasLegend = False
 Probe = PlotHost.ChartTitle.Text & ":" & CStr(PlotHost.HasLegend) & ":" & CStr(PlotHost.Tick())
End Function
