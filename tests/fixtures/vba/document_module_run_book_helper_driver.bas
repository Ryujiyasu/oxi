Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
 Probe = CStr(ThisWorkbook.Helper(7))
End Function
Private Sub Unused()
 Dim value As Variant
 Dim target As Object
 value = 1
End Sub
