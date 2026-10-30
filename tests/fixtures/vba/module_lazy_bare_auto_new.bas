Attribute VB_Name = "DriverModule"
Option Explicit
Public initialized As Long
Public Function Probe() As String
On Error Resume Next
Probe = CStr(initialized) & ":" & CStr(holder.Value) & ":" & CStr(initialized)
End Function
Attribute VB_Name = "CounterModule"
Public holder As New Sample
VERSION 1.0 CLASS
Attribute VB_Name = "Sample"
Private Sub Class_Initialize()
initialized = initialized + 1
End Sub
Public Property Get Value() As Long
Value = 7
End Property
