Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = HelperA.Secret(Nothing)
End Sub
Attribute VB_Name = "HelperA"
Private Function Secret(ByVal required As Long) As Long
 Secret = required
End Function
