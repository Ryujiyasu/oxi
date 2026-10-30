Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim target As Receiver
  Dim value As Variant
  value = target.Item
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Property Let Item(ByVal value As Variant)
End Property
