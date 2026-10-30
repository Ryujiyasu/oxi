Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim number As Long
number = 1
OtherModule.Item(number) = 9
Probe = CStr(number) & ":" & CStr(OtherModule.number)
End Function
Attribute VB_Name = "OtherModule"
Option Explicit
Public number As Long
Public Property Let Item(ByRef index As Long, ByVal value As Long)
number = 50
index = index + 10
End Property
