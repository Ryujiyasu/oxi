Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = "ran"
End Function
Private Sub Hidden()
Dim value As Office.MsoTriState
value = msoTrue
End Sub
