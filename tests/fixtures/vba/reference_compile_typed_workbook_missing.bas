Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim target As Excel.Workbook
Dim value As Variant
value = target.NoSuchMember
End Sub
