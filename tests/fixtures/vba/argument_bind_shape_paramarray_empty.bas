Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = Gather()
End Sub
Private Function Gather(ParamArray args() As Variant) As Variant
  Gather = 0
End Function
