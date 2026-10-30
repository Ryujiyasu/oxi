Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Values
Item = 2147483647
NextItem
End Enum
Public Function Probe() As String
On Error Resume Next
Probe = CStr(Item) & ":" & CStr(NextItem) & ":" & CStr(VarType(NextItem))
End Function
