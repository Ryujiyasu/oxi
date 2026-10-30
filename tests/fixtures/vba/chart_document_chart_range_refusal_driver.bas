Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim target As Object, value As Object
 Set target = PlotHost
 On Error Resume Next
 Set value = target.Range("A1")
 Probe = CStr(Err.Number) & ":" & TypeName(value)
End Function
