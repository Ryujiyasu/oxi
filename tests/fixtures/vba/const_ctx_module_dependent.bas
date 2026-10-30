Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim value(SizesModule.Derived) As Long
Probe = CStr(UBound(value))
End Function
Attribute VB_Name = "SizesModule"
Option Explicit
Private Const seed = 1
Public Const Low = -1
Public Const High = 2
Public Const Derived = seed + 1
Private Const Hidden = 2
