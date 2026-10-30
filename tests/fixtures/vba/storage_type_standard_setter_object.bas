Attribute VB_Name = "DriverModule"
Option Explicit
Public held As WorkerClass
Public Function Probe() As String
Dim other As New WorkerClass
Dim number As Long
number = 1
Set Item(number) = other
Probe = CStr(number) & ":" & CStr(held Is other)
End Function
Public Property Set Item(ByRef index As Long, ByVal value As WorkerClass)
index = index + 10
Set held = value
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "WorkerClass"
Public number As Long
