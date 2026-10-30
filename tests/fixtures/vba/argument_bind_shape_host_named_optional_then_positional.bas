Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = Application.WorksheetFunction.SumIf(Arg3:=Nothing, Nothing, 1)
End Sub
