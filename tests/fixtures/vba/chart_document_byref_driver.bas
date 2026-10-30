Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim value As Long, target As Object
 value = 7
 Set target = PlotHost
 target.Bump value
 Probe = CStr(value) & ":" & CStr(PlotHost.Tick())
End Function
