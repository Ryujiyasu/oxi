Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim s As Object
 Set s = Ledger
 Dim n As Long
 n = 7
 s.Bump n
 Probe = CStr(n)
End Function
