Attribute VB_Name = "DriverModule"
Option Explicit
Public Function TargetChart() As Object
 Set TargetChart = PlotHost
End Function
Public Sub Change()
 Dim target As Object
 Set target = TargetChart()
 target.HasTitle = True
 target.ChartTitle.Text = "売上 & <利益>" & vbLf & "2026"
End Sub
Public Function Probe() As String
 Dim target As Object, title As String, position As String
 Set target = TargetChart()
 title = "[off]"
 position = "[off]"
 If target.HasTitle Then title = target.ChartTitle.Text
 If target.HasLegend Then position = CStr(target.Legend.position)
 Probe = CStr(target.HasTitle) & "|" & title & "|" & CStr(target.HasLegend) & "|" & position
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArchiveHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlotHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

