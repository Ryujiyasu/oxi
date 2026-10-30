Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = "ran"
End Function
Private Sub Hidden()
Dim value As XlCalculation
value = xlCalculationManual
End Sub
