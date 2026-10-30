Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = GatherRequired()
End Sub
Private Function GatherRequired(required As Variant, ParamArray args() As Variant) As Variant
  GatherRequired = 0
End Function
