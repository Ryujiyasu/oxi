Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
value = "x"
Probe = CStr(Len(value)) & ":<" & value & ">"
End Function
Attribute VB_Name = "CounterModule"
Public value As String * 3
