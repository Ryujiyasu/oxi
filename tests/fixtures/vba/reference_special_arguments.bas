Attribute VB_Name = "DriverModule"
Private held As CounterClass
Public Function Probe() As String
If held Is Nothing Then Set held = New CounterClass
Probe = TypeName(PassValue(Date)) & ":" & held.WhoAmI()
End Function
Private Function PassValue(ByRef value As Variant) As Variant
PassValue = value
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "CounterClass"
Public Function WhoAmI() As String
WhoAmI = TypeName(Identity(Me))
End Function
Private Function Identity(ByRef value As Object) As Object
Set Identity = value
End Function
