Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Values
Item = CInt(2)
End Enum
Public Function Probe() As String
On Error Resume Next
Dim arr(Item) As Long
Probe = CStr(Item) & ":" & CStr(UBound(arr))
End Function
