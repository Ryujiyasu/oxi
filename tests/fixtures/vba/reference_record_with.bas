Attribute VB_Name = "DriverModule"
Option Explicit
Public calls As Long
Public Function NextIndex() As Long
calls = calls + 1
NextIndex = 1
End Function
Public Function Probe() As String
With LeftModule.entries(NextIndex())
.amount = .amount + 1
.label = "abcdef"
RightModule.Change .amount
End With
Probe = LeftModule.entries(1).amount & ":" & LeftModule.entries(1).label & ":" & calls & ":" & TypeName(LeftModule.entries(1).amount)
End Function
Attribute VB_Name = "LeftModule"
Option Explicit
Public Type Entry
amount As Long
label As String * 4
End Type
Public entries(1 To 1) As Entry
Attribute VB_Name = "RightModule"
Option Explicit
Public Sub Change(ByRef value As Long)
value = value + 1
End Sub
