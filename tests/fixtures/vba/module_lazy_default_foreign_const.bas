Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe(Optional ByVal item As Long = seed) As String
Probe = CStr(item)
End Function
Attribute VB_Name = "ValuesModule"
Public Const seed = 2
