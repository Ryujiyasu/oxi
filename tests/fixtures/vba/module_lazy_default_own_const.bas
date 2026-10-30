Attribute VB_Name = "DriverModule"
Option Explicit
Public Const seed = 2
Public Function Probe(Optional ByVal item As Long = seed) As String
Probe = CStr(item)
End Function
