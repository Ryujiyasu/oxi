Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As Object
 Set s = Ledger
 On Error Resume Next
 Dim n As Long
 n = s.Helper(7)
 Probe = CStr(Err.Number) & ":" & CStr(n)
End Function
