Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As LongPtr = 1E+20
Probe = "ran"
End Function
