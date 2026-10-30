Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim target As Receiver
  Dim value As Variant
  value = target.Item(1)
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Property Let Item(ByVal index As Variant, ByVal value As Variant)
End Property
