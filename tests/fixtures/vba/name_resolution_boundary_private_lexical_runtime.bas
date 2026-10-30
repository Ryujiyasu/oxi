Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Dim target As New Receiver
 Probe = target.RunInner()
End Function
Private Sub Unused()
  Dim value As Variant
  
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Function RunInner() As String
 RunInner = CStr(Hidden(1))
End Function
Private Function Hidden(ByVal required As Long) As Long
 Hidden = required
End Function
