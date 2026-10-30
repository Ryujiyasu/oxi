Attribute VB_Name = "DriverModule"
Option Explicit
Public seen As String
Public Function Probe() As String
Dim worker As New WorkerClass
worker.number = 3
On Error Resume Next
Call Change (worker)
Probe = CStr(Err.Number) & ":" & seen & ":" & CStr(worker.number)
End Function
Private Sub Change(ByRef value As Variant)
seen = TypeName(value)
value = 7
End Sub
Private Function Consume(ByRef value As Variant) As String
Consume = TypeName(value)
value = 7
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public number As Long
Public Property Let Value(ByVal value As Long)
Attribute Value.VB_UserMemId = 0
number = value
End Property
Public Property Get Value() As Long
Value = number
End Property
