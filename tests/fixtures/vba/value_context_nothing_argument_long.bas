Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Probe = Consume(Nothing)
End Function
Private Function Consume(ByVal value As Long) As String
Consume = CStr(value)
End Function
