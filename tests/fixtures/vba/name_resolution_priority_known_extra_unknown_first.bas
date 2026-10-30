Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Sub Unused()
 Dim value As Variant
 value = Known(MissingVariable, 1)
End Sub
Private Function Known(ByVal required As Long) As Long
 Known = required
End Function
