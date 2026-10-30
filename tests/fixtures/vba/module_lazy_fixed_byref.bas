Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Bump value
Probe = CStr(Len(value)) & ":<" & value & ">"
End Function
Private Sub Bump(ByRef arg As String)
arg = "abcde"
End Sub
Attribute VB_Name = "CounterModule"
Public value As String * 3
