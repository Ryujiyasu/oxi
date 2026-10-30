Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
Dim value As Variant
Set value = Nothing
Probe = CStr(value Is Nothing)
End Function
