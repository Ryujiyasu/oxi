Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim value As Variant
Dim worker As New WorkerClass
worker.number = 3
On Error Resume Next
value = (worker)
Probe = CStr(Err.Number) & ":" & TypeName(value)
If IsObject(value) Then
If Not value Is Nothing Then Probe = Probe & ":" & CStr(value.number)
End If
End Function
Private Function Consume(ByVal item As Object) As Object
Set Consume = item
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
