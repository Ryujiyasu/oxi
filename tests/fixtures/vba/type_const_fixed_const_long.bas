Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As String * 3 = "abcdef"
Probe = CStr(Len(value)) & ":<" & value & ">"
End Function
