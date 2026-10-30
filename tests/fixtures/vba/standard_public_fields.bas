Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
LeftModule.Count = LeftModule.Count + 1
RightModule.Count = RightModule.Count + 10
Probe = LeftModule.ReadCount() & "|" & RightModule.ReadCount()
End Function

Attribute VB_Name = "LeftModule"
Option Explicit
Public Count As Long
Public Function ReadCount() As Long
ReadCount = Count
End Function

Attribute VB_Name = "RightModule"
Option Explicit
Public Count As Long
Public Function ReadCount() As Long
ReadCount = Count
End Function
