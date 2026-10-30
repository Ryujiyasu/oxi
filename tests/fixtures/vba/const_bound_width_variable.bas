Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim size As Long
Dim value As String * size
Probe = CStr(Len(value))
End Function
