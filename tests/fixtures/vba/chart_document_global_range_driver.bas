Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Dim target As Object
 On Error Resume Next
 Set target = Range("A1")
 Probe = CStr(Err.Number) & ":" & TypeName(target)
End Function
