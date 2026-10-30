Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim number As Long
number = 1
Item(number) = 9
Probe = CStr(number)
End Function
Private Property Let Item(ByRef index As Long, ByVal value As Long)
index = index + 10
End Property
