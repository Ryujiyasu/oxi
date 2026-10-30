Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Probe = LeftModule.Inspect() & "|" & RightModule.Inspect()
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Option Compare Text
Option Base 1
Public Function Inspect() As String
Dim values(2) As Long
Inspect = CStr("a" = "A") & ":" & LBound(values)
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Option Compare Binary
Option Base 0
Public Function Inspect() As String
Dim values(2) As Long
Inspect = CStr("a" = "A") & ":" & LBound(values)
End Function
