Attribute VB_Name = "DriverModule"
Option Explicit
Public Function TargetChart() As Object
 Set TargetChart = PlotHost
End Function
Public Sub Change()
 Dim target As Object
 Set target = TargetChart()
 target.HasTitle = False
 target.HasTitle = True
End Sub
Public Function Probe() As String
 Dim target As Object, title As String, position As String
 Set target = TargetChart()
 title = "[off]"
 position = "[off]"
 If target.HasTitle Then title = target.ChartTitle.Text
 If target.HasLegend Then position = CStr(target.Legend.Position)
 Probe = CStr(target.HasTitle) & "|" & title & "|" & CStr(target.HasLegend) & "|" & position
End Function
