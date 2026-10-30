Attribute VB_Name = "DriverModule"
Option Explicit
Public counter As Long
Public Function Probe() As String
counter = counter + 1
Probe = CStr(counter)
End Function
Private Sub Hidden()
missing = 1
End Sub
