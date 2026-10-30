Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim target As Receiver
  Dim value As Variant
  target.Item(1) = 2
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Property Get Item(ByVal index As Variant) As Variant
  Item = index
End Property
Public Property Let Item(ByVal index As Variant, ByVal value As Variant)
End Property
