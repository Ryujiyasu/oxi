Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim target As Object
 Set target = New PlotHost
 Probe = TypeName(target)
End Function
