' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102143
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Private Function TypedShape(ByVal value As Variant) As String
    Dim dimension As Long, lower As Long, upper As Long, failure As Long
    Dim answer As String, member As Variant
    If IsObject(value) Then
        TypedShape = "object:" & TypeName(value) & ":" & value.Address(False, False)
    ElseIf IsNull(value) Then
        TypedShape = "1:null"
    ElseIf IsArray(value) Then
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
Public Function ErrorDomain(ByVal first As Long, ByVal last As Long, ByVal route As Long, ByVal arrayMode As Long) As String
    Dim target As Object, argument As Variant, value As Variant
    Dim code As Long, failed As Long, previous As Long, report As String
    On Error GoTo Unexpected
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    If route = 0 Then Set target = Application.WorksheetFunction Else Set target = Application
    For code = first To last
        On Error Resume Next
        Err.Clear
        argument = CVErr(code)
        failed = Err.Number
        On Error GoTo Unexpected
        report = report & CStr(code) & ":ctor=" & CStr(failed)
        If failed = 0 Then
            If arrayMode = 1 Then argument = Array(argument)
            On Error Resume Next
            Err.Clear
            value = Empty
            value = CallByName(target, "Phonetic", VbMethod, argument)
            failed = Err.Number
            On Error GoTo Unexpected
            report = report & ":call=" & CStr(failed)
            If failed = 0 Then report = report & ":" & TypedShape(value)
        End If
        report = report & "|"
    Next code
    ErrorDomain = report
    Application.Calculation = previous
    Exit Function
Unexpected:
    failed = Err.Number: ErrorDomain = "unexpected:" & CStr(failed) & ":" & report
    On Error Resume Next
    Application.Calculation = previous
End Function
Public Function MemberCase(ByVal kind As Long, ByVal route As Long, ByVal spelling As Long) As String
    Dim target As Object, argument As Variant, value As Variant, member As String
    Dim failed As Long, previous As Long, inputDescription As String
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    If route = 0 Then Set target = Application.WorksheetFunction Else Set target = Application
    If spelling = 0 Then member = "Error_Type"
    If spelling = 1 Then member = "ErrorType"
    If spelling = 2 Then member = "ERROR.TYPE"
    If kind = 0 Then argument = 1
    If kind = 1 Then argument = CVErr(2042)
    If kind = 2 Then
        Range("D4").Value2 = CVErr(2042)
        Set argument = Range("D4")
    End If
    inputDescription = TypedShape(argument)
    value = CallByName(target, member, VbMethod, argument)
    MemberCase = "input=" & inputDescription & "|result=" & TypedShape(value)
    Range("D4").ClearContents
    Application.Calculation = previous
    Exit Function
Failed:
    failed = Err.Number: MemberCase = "input=" & inputDescription & "|fault=" & CStr(failed)
    On Error Resume Next
    Range("D4").ClearContents
    Application.Calculation = previous
End Function
