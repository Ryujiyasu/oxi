Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim refs(0) As Object
Set refs(0) = Range("A1")
On Error Resume Next
refs(0) = 7
Probe = CStr(Err.Number) & ":" & TypeName(refs(0)) & ":" & CStr(Range("A1").Value)
End Function
