Attribute VB_Name = "DriverModule"
Option Explicit
Public Function Probe() As String
Dim number As LongLong
number = 1
Probe = Inspect(number)
End Function
Private Function Inspect(ByRef value As Variant) As String
Dim kind As Long
Dim label As String
Dim copied As Variant
On Error Resume Next
kind = VarType(value)
Inspect = kind & ":" & Err.Number
Err.Clear
label = TypeName(value)
Inspect = Inspect & ":" & label & ":" & Err.Number
Err.Clear
copied = value
Inspect = Inspect & ":" & TypeName(copied) & ":" & Err.Number
Err.Clear
value = 2.6
Inspect = Inspect & ":" & Err.Number
End Function
