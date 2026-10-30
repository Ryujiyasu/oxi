Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Sub Unused()
 Dim value As Variant
 Dim target As Object
 Dim s As Ledger
 Set s = Ledger
 value = s.Helper()
End Sub
