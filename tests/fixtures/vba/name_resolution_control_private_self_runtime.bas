Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Dim target As New Receiver
  Probe = target.RunInner()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Function RunInner() As String
  RunInner = CStr(Me.Hidden(1))
End Function
Private Function Hidden(ByVal required As Long) As Long
  Hidden = required
End Function
