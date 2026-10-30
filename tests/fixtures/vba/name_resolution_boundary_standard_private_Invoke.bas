Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = DriverModule.Invoke()
End Sub
Private Function Invoke() As Long
 Invoke = 1
End Function
