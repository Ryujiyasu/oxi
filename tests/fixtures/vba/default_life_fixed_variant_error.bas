Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim text As String * 4
text = "abcd"
On Error Resume Next
Change text
Probe = CStr(Err.Number) & ":[" & text & "]"
End Function
Private Sub Change(ByRef value As Variant)
value = "xyz"
Err.Raise 513
End Sub
