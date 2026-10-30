' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102139
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Private Function TypedShape(ByVal value As Variant) As String
    Dim dimension As Long, lower As Long, upper As Long, failure As Long
    Dim answer As String, member As Variant
    If IsArray(value) Then
        answer = CStr(VarType(value)) & ":bounds["
        For dimension = 1 To 60
            On Error Resume Next
            Err.Clear
            lower = LBound(value, dimension)
            failure = Err.Number
            If failure = 0 Then
                upper = UBound(value, dimension)
                failure = Err.Number
            End If
            On Error GoTo 0
            If failure <> 0 Then Exit For
            answer = answer & CStr(lower) & ":" & CStr(upper) & ";"
        Next dimension
        answer = answer & "]values["
        For Each member In value
            answer = answer & TypedShape(member) & ";"
        Next member
        TypedShape = answer & "]"
    ElseIf IsError(value) Then
        TypedShape = CStr(VarType(value)) & ":errorCode=" & CStr(CLng(value))
    ElseIf IsObject(value) Then
        TypedShape = CStr(VarType(value)) & ":" & TypeName(value) & ":" & value.Address(False, False)
    Else
        TypedShape = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function
Public Function ShapeCase(ByVal kind As Long, ByVal route As Long) As String
    Dim previous As Long, argument As Variant, value As Variant, target As Object
    Dim strings() As String, x As Long, y As Long, z As Long
    Dim inputDescription As String, failure As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("D4:D5").Value2 = "ABC"
    Range("D5").Value2 = "DEF"
    Select Case kind
        Case 0: argument = "ABC"
        Case 1: argument = Array("ABC", "DEF")
        Case 2
            ReDim argument(5 To 6)
            argument(5) = "ABC": argument(6) = "DEF"
        Case 3
            ReDim argument(2 To 3, -1 To 1)
            For x = 2 To 3
                For y = -1 To 1
                    argument(x, y) = CStr(x) & ":" & CStr(y)
                Next y
            Next x
        Case 4
            ReDim argument(0 To 1, 0 To 1, 0 To 1)
            For x = 0 To 1
                For y = 0 To 1
                    For z = 0 To 1
                        argument(x, y, z) = CStr(x) & ":" & CStr(y) & ":" & CStr(z)
                    Next z
                Next y
            Next x
        Case 5: Set argument = Range("D4:D5")
        Case 6: argument = Range("D4:D5").Value2
        Case 7: argument = Array()
        Case 8
            ReDim strings(5 To 6)
            strings(5) = "ABC": strings(6) = "DEF"
            argument = strings
        Case 9: argument = CVErr(2042)
        Case 10: argument = Empty
        Case 11: argument = Application.Evaluate("{""ABC"",""DEF""}")
    End Select
    inputDescription = TypedShape(argument)
    If route = 0 Then
        Set target = Application.WorksheetFunction
    Else
        Set target = Application
    End If
    value = CallByName(target, "Phonetic", VbMethod, argument)
    ShapeCase = "input=" & inputDescription & "|result=" & TypedShape(value)
    Range("D4:D5").ClearContents
    Application.Calculation = previous
    Exit Function
Failed:
    failure = Err.Number
    ShapeCase = "input=" & inputDescription & "|fault=" & CStr(failure)
    On Error Resume Next
    Range("D4:D5").ClearContents
    Application.Calculation = previous
End Function
