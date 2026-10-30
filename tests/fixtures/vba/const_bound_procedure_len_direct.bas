Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = CStr(Len("a"))
End Function
Private Function Len(ByVal value As String) As Long
Len = 2
End Function
