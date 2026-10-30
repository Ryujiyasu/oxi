Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim count As Long
count = count + 1
Probe = CStr(count)
End Function
Attribute VB_Name = "BadModule"
Public Const count As Byte = 300
