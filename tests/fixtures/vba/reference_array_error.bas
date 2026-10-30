Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
LeftModule.items(1) = 1
On Error Resume Next
RightModule.Change LeftModule.items(1)
Probe = Err.Number & ":" & LeftModule.items(1)
End Function
Attribute VB_Name = "LeftModule"
Option Explicit
Public items(1 To 1) As Long
Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Change(ByRef value As Long)
value = value + 1
Err.Raise 513
End Sub
