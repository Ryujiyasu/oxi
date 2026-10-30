Attribute VB_Name = "DriverModule"
Option Explicit
Private Const size = 3
Public Function Probe() As String
On Error Resume Next
Dim value As String * size
Probe = CStr(Len(value))
End Function
