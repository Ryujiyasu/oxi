Attribute VB_Name = "DriverModule"
Option Explicit
Public touched As Long
Public holder As New Sample
Public Function Probe() As String
Probe = CStr(touched)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Sample"
Private Sub Class_Initialize()
DriverModule.touched = DriverModule.touched + 1
End Sub
