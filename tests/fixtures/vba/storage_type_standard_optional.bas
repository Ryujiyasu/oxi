Attribute VB_Name = "DriverModule"
Option Explicit
Public Log As String
Public Property Let Item(Optional ByVal index As Long = 3, ByVal value As Long)
Log = CStr(index) & ":" & CStr(value)
End Property
Public Function Probe() As String
Item() = 9
Probe = Log
End Function
