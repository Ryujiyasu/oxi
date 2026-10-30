Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Sub Unused()
 Dim value As Variant
 value = Known(NoSuchArg:=MissingVariable)
End Sub
Private Function Known(ByVal required As Long) As Long
 Known = required
End Function
