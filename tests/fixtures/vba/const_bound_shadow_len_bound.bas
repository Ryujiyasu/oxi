Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value(Len("a")) As Long
Probe = CStr(UBound(value))
End Function
Private Function Len(ByVal value As String) As Long
Len = 2
End Function
