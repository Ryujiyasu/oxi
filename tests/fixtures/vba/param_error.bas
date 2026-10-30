Attribute VB_Name = "DriverModule"
Option Explicit
Public number As Long
Public Function Probe() As String
number = 1
On Error Resume Next
Change number
Probe = Err.Number & ":" & number
End Function
Private Sub Change(ParamArray values() As Variant)
values(0) = values(0) + 1
Err.Raise 513
End Sub
