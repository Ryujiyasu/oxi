Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = DriverModule.Hidden()
End Sub
Private Function Hidden(ByVal required As Long) As Long
 Hidden = required
End Function
