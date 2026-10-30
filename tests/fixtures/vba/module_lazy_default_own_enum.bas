Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Values
SeedValue = 2
End Enum
Public Function Probe(Optional ByVal item As Long = SeedValue) As String
Probe = CStr(item)
End Function
