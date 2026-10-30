Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = Application.WorksheetFunction.SumIf(Arg1:=Nothing, Arg2:=1, Arg1:=Nothing)
End Sub
