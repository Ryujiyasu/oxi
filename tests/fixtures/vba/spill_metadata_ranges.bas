' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102115
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

Public Function FlowCase(ByVal targetKind As Long, ByVal adjacent As Long) As String
    Dim previous As Long, stage As String, selected As Range, address As String
    Dim answer As String, code As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    stage = "setup"
    Range("M1:R4").ClearContents
    Range("N1").Formula2 = "=SEQUENCE(2,2)"
    If adjacent <> 0 Then Range("P1").Formula2 = "=SEQUENCE(2,2,11)"
    Application.CalculateFull
    Select Case targetKind
        Case 0: address = "N1"
        Case 1: address = "N2"
        Case 2: address = "N1:O2"
        Case 3: address = "N1:O3"
        Case 4: address = "M1:O2"
        Case 5: address = "N1:Q2"
        Case 6: address = "O1:O2"
        Case 7: address = "P1:Q2"
        Case 8: address = "N3"
        Case 9: address = "N1,P1"
    End Select
    stage = "facts"
    Set selected = Range(address)
    answer = SpillFacts(selected)
    stage = "cleanup"
    Range("M1:R4").ClearContents
    Application.Calculation = previous
    FlowCase = answer
    Exit Function
Failed:
    code = Err.Number
    FlowCase = "error@" & stage & ":" & CStr(code)
    On Error Resume Next
    Range("M1:R4").ClearContents
    Application.Calculation = previous
End Function
