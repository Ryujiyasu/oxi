Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
  Probe = "ran"
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Function RunInner() As String
  RunInner = CStr(Me.Hidden())
End Function
Private Function Hidden(ByVal required As Long) As Long
  Hidden = required
End Function
