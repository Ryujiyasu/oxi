Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Call Known(1)
End Sub
Private Sub Known(ByVal required As Long)
End Sub
Private Function Result(ByVal required As Long) As Long
  Result = required
End Function
