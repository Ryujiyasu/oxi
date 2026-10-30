Attribute VB_Name = "DriverModule"
Option Explicit

Public Function Probe() As String
On Error Resume Next
Probe = Consume(Nothing)
End Function
Private Function Consume(ByVal value As Object) As String
Consume = CStr(value Is Nothing)
End Function
