Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = "ran"
End Function
Private Sub Hidden()
CallByName 1, "x", 2
End Sub
