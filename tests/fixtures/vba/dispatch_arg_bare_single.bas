Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Function Probe() As String
On Error Resume Next
Change(1)
Probe = CStr(Err.Number) & ":" & Log
End Function
Private Sub Change(ByVal first As Long, Optional ByVal second As Long = 2)
Log = "ran"
End Sub
