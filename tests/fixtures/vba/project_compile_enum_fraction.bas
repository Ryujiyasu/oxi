Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Values
Item = 1.5
End Enum
Public Function Probe() As String
On Error Resume Next
Probe = CStr(Item) & ":" & CStr(VarType(Item))
End Function
