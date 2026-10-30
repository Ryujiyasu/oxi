Attribute VB_Name = "DriverModule"
Option Explicit
Private Sub Change(ByVal value As FirstClass)
End Sub
Private Function MakeValue(ByVal source As Object) As FirstClass
Set MakeValue = source
End Function
Public Function Probe() As String
Dim source As Object, target As FirstClass
Set source = New OtherClass
On Error Resume Next
Set target = MakeValue(source)
Probe = CStr(Err.Number) & ":" & CStr(target Is Nothing)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Option Explicit
Public number As Long
VERSION 1.0 CLASS
Attribute VB_Name = "FirstClass"
Public number As Long
VERSION 1.0 CLASS
Attribute VB_Name = "OtherClass"
Public number As Long
