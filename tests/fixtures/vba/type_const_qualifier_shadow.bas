Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
On Error Resume Next
Dim SizesModule As Object
Dim value(SizesModule.High) As Long
Probe = "ran"
End Function
Attribute VB_Name = "SizesModule"
Public Const High = 2
