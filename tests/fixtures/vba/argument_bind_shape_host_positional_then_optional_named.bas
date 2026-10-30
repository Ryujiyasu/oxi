Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = Application.WorksheetFunction.SumIf(Nothing, 1, Arg3:=Nothing)
End Sub
