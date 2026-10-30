Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = TypeName(Worksheets) & "|" & TypeName(Sheets) & "|" & TypeName(Charts) & "|" & TypeName(ActiveSheet) & "|" & TypeName(PlotHost) & "|" & CStr(Worksheets(1).Index) & "|" & CStr(PlotHost.Index) & "|" & CStr(PlotHost Is Sheets(1)) & "|" & CStr(PlotHost Is Charts(1))
End Function
