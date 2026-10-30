Attribute VB_Name = "DriverModule"
Option Explicit
Public seen As String
Public Function Probe() As String
Dim text As String * 4
text = "abcd"
Change text, text
Probe = seen & ":[" & text & "]"
End Function
Private Sub Change(ByRef first As String, ByRef second As String)
first = "xyz"
seen = first & ":" & second
End Sub
