Attribute VB_Name = "DriverModule"
Option Explicit
Public Function ReadBoth() As String
ReadBoth = CStr(LeftModule.NextCount()) & "|" & CStr(RightModule.NextCount())
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public Function NextCount() As Long
Static stored As Long
stored = stored + 1
NextCount = stored
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Public Function NextCount() As Long
Static stored As Long
stored = stored + 1
NextCount = stored
End Function
