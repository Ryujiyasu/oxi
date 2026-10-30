Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = Mirror(1, 2, third:=3)
End Sub
Private Function Mirror(first As Variant, second As Variant, Optional third As Variant) As Variant
  Mirror = first
End Function
