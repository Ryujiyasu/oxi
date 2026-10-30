Attribute VB_Name = "DriverModule"
Option Explicit
Public touched As Long
Public Function Probe() As String
On Error Resume Next
Probe = CStr(touched)
End Function
Public Function Touch() As Long
touched = touched + 1
Touch = 2
End Function
Attribute VB_Name = "BadModule"
Public Const broken = DriverModule.Touch()
