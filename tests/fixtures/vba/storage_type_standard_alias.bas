Attribute VB_Name = "DriverModule"
Option Explicit
Public Property Let Item(ByRef first As Long, ByRef second As Long, ByVal value As Long)
first = first + 1
second = second + first
End Property
Public Function Probe() As String
Dim number As Long
number = 1
Item(number, number) = 9
Probe = CStr(number)
End Function
