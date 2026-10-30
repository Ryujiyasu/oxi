Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim target As Worksheet
Dim value As Variant
With target
value = .NoSuchMember
End With
End Sub
