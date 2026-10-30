Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim result As Variant
#If False Then
result = Len(CLng(123))
#End If
Probe = "after"
End Function
