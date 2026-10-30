Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function Probe() As String
On Error Resume Next
Call Change(1, 2)
Probe = CStr(Err.Number) & ":" & Log
End Function
Private Sub Change(ByVal first As Long, ByVal second As Long)
Log = "ran"
End Sub
