Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
value = 5
Probe = TypeName(value) & ":" & CStr(value < "10")
End Function
Attribute VB_Name = "CounterModule"
Public value As Variant
