Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Values
Item = Date()
End Enum
Public Function Probe() As String
On Error Resume Next
Probe = CStr(Item)
End Function
