Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Sub Unused()
 Dim value As Variant
 Dim target As Object
 With Ledger
 value = .Helper()
 End With
End Sub
