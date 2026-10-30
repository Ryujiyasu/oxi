Attribute VB_Name = "DriverModule"
Option Explicit
Public calls As Long
Public Function NextIndex() As Long
calls = calls + 1
NextIndex = 1
End Function
Public Function Probe() As String
LeftModule.items(1) = 1
RightModule.Change (LeftModule.items(NextIndex()))
RightModule.ChangeValue LeftModule.items(NextIndex())
RightModule.Change LeftModule.items(NextIndex())
Probe = LeftModule.items(1) & ":" & calls
End Function
Attribute VB_Name = "LeftModule"
Option Explicit
Public items(1 To 1) As Long
Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Change(ByRef value As Long)
value = value + 1
End Sub
Public Sub ChangeValue(ByVal value As Long)
value = value + 10
End Sub
