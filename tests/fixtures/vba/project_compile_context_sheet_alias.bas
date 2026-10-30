Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = "ran"
End Function
Private Sub Hidden()
Sheet1.Range("A1").Value = 1
End Sub
