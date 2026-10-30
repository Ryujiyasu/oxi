Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = "ran"
End Function
Private Sub Unused()
 Dim value As Variant
 Dim target As Object
 Dim b As ThisWorkbook
 Set b = ThisWorkbook
 value = b.Helper(7)
End Sub
