Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Values
Item = GetValue()
End Enum
Public Function Probe() As String
On Error Resume Next
Probe = CStr(Item)
End Function
Private Function GetValue() As Long
GetValue = 2
End Function
