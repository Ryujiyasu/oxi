Option Explicit
Private moduleValue As Long
Private moduleObject As Object
Private Function Counter() As Long
    Static stored As Long
    stored = stored + 1
    Counter = stored
End Function
Public Function ReadGlobals() As String
    Dim reference As String
    If moduleObject Is Nothing Then
        reference = "none"
    Else
        reference = moduleObject.Address(False, False) & ":" & CStr(moduleObject.Value)
    End If
    ReadGlobals = CStr(moduleValue) & "|" & CStr(moduleObject Is Nothing) & "|" & CStr(Counter()) & "|" & reference
End Function
Public Function Prime(ByVal exitKind As Long) As String
    moduleValue = 123
    Set moduleObject = Range("A1")
    Range("A1").Value = 5
    Range("D1").Value = ReadGlobals()
    Select Case exitKind
    Case 1
        Error 11
    Case 2
        Error 40040
    Case 3
        On Error Resume Next
        Error 11
    Case 4
        Err.Raise 513, "own", "failure"
    Case 6
        End
    Case 5
        Dim value As Variant
        value = Range("A1").Offset(-1, 0).Value
    End Select
    Prime = ReadGlobals()
End Function
