Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim result As Variant
On Error Resume Next
If False Then result = Len(CLng(123))
Probe = "after"
End Function
