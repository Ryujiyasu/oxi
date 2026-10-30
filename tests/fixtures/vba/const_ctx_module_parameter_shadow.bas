Attribute VB_Name = "DriverModule"
Option Explicit
Public Const size = 2
Public Function Probe() As String
On Error Resume Next
Helper 1
Probe = "ran"
End Function
Private Sub Helper(ByVal size As Long)
Dim value(size) As Long
End Sub
