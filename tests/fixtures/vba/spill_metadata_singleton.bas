' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102119
End Function

Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function

Private Function SpillFacts(ByVal target As Range) As String
    Dim v As Variant, held As Range, answer As String, code As Long
    On Error Resume Next
    Err.Clear
    v = target.HasSpill
    code = Err.Number
    If code <> 0 Then
        answer = "has:error" & CStr(code)
    ElseIf IsNull(v) Then
        answer = "has:" & CStr(VarType(v)) & ":Null"
    Else
        answer = "has:" & CStr(VarType(v)) & ":" & CStr(v)
    End If
    Err.Clear
    Set held = Nothing
    Set held = target.SpillParent
    code = Err.Number
    If code <> 0 Then
        answer = answer & "|parent:error" & CStr(code)
    ElseIf held Is Nothing Then
        answer = answer & "|parent:Nothing"
    Else
        answer = answer & "|parent:" & held.Address
    End If
    Err.Clear
    Set held = Nothing
    Set held = target.SpillingToRange
    code = Err.Number
    If code <> 0 Then
        answer = answer & "|range:error" & CStr(code)
    ElseIf held Is Nothing Then
        answer = answer & "|range:Nothing"
    Else
        answer = answer & "|range:" & held.Address
    End If
    Err.Clear
    v = target.HasFormula
    code = Err.Number
    If code <> 0 Then
        answer = answer & "|formula:error" & CStr(code)
    ElseIf IsNull(v) Then
        answer = answer & "|formula:" & CStr(VarType(v)) & ":Null"
    Else
        answer = answer & "|formula:" & CStr(VarType(v)) & ":" & CStr(v)
    End If
    SpillFacts = answer
End Function

Private Function SpillSuffix() As String
    Dim held As Range, code As Long
    On Error Resume Next
    Err.Clear
    Set held = Range("N1#")
    code = Err.Number
    If code <> 0 Then
        SpillSuffix = "error" & CStr(code)
    ElseIf held Is Nothing Then
        SpillSuffix = "Nothing"
    Else
        SpillSuffix = held.Address
    End If
End Function

Public Function FlowCase(ByVal kind As Long, ByVal site As Long) As String
    Dim previous As Long, expression As String, answer As String, code As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:R4").ClearContents
    Select Case kind
        Case 0: expression = "=SEQUENCE(1)"
        Case 1: expression = "=SEQUENCE(1,1)"
        Case 2: expression = "=SEQUENCE(1,2)"
        Case 3: expression = "=SEQUENCE(2,1)"
        Case 4: expression = "=1"
        Case 5: expression = "=IF(TRUE,SEQUENCE(1),0)"
        Case 6: expression = "=TRANSPOSE(SEQUENCE(1))"
        Case 7: expression = "=FILTER(SEQUENCE(2),SEQUENCE(2)=1)"
        Case 8: expression = "=SORT(SEQUENCE(1))"
        Case 9: expression = "=IFERROR(SEQUENCE(1),0)"
    End Select
    Range("N1").Formula2 = expression
    Application.CalculateFull
    If site = 0 Then
        answer = SpillFacts(Range("N1"))
    Else
        answer = SpillFacts(Range("N2"))
    End If
    FlowCase = answer & "|suffix:" & SpillSuffix()
    Range("M1:R4").ClearContents
    Application.Calculation = previous
    Exit Function
Failed:
    code = Err.Number
    FlowCase = "error@singleton:" & CStr(code)
    On Error Resume Next
    Range("M1:R4").ClearContents
    Application.Calculation = previous
End Function
