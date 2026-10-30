Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 With Ledger
 Probe = CStr(.Helper(required:=7))
 End With
End Function
