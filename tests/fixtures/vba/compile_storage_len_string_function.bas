Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String

On Error Resume Next
Dim result As Variant
result = Len(GetText())
Probe = Err.Number & ":" & TypeName(result)
If Not IsNull(result) Then Probe = Probe & ":" & CStr(result)
End Function
Private Function GetText() As String
GetText = "123"
End Function
