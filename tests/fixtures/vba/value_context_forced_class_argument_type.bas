Attribute VB_Name = "DriverModule"
Option Explicit
Public seen As String
Public Function Probe() As String
Dim worker As New WorkerClass
worker.number = 3
On Error Resume Next
Change (worker)
Probe = CStr(Err.Number) & ":" & seen
End Function
Private Sub Change(ByVal value As Variant)
seen = TypeName(value)
End Sub
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
