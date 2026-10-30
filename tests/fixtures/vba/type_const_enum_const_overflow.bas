Attribute VB_Name = "DriverModule"
Option Explicit
Public Enum Choice
One = 1
End Enum
Public Function Probe() As String
On Error Resume Next
Const value As Choice = 2147483648#
Probe = "ran"
End Function
