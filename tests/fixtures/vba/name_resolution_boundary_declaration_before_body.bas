Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
Private Sub Unused()
  Dim value As Variant
  value = MissingVariable
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Function Invoke() As Long
 Invoke = 1
End Function
