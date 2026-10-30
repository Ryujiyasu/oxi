Attribute VB_Name = "DriverModule"
Option Explicit
Public trace As String
Public Function Probe() As String
ReDim LeftModule.items(1 To 1)
LeftModule.items(1) = 1
RightModule.Change LeftModule.items(1)
Probe = trace & ":" & LeftModule.items(1) & ":" & UBound(LeftModule.items)
End Function
Attribute VB_Name = "LeftModule"
Option Explicit
Public items() As Long
Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Change(ByRef value As Long)
On Error Resume Next
ReDim LeftModule.items(1 To 2)
trace = CStr(Err.Number)
Err.Clear
value = value + 1
End Sub
