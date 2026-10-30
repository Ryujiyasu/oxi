Attribute VB_Name = "DriverModule"
Option Explicit
Private Property Let Item(ByRef index As Long, ByVal value As Long)
index = 11
Err.Raise 513
End Property
Public Function Probe() As String
Dim number As Long
number = 1
On Error Resume Next
Item(number) = 9
Probe = CStr(Err.Number) & ":" & CStr(number)
End Function
