Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(VBA.Len("ab")) As Long
Probe = CStr(UBound(value))
End Function
