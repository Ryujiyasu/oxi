Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Const value As String * 3 = "x"
Dim arr(Len(value)) As Long
Probe = CStr(UBound(arr))
End Function
