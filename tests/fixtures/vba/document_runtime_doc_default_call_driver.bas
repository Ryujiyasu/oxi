Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 On Error Resume Next
 Dim s As Object
 Set s = Ledger("A1")
 Probe = CStr(Err.Number) & ":" & TypeName(s)
End Function
