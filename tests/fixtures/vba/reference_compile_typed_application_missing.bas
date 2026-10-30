Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim target As Excel.Application
Dim value As Variant
value = target.NoSuchMember
End Sub
