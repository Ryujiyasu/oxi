' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102118
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

Public Function CacheStage(ByVal mode As Long) As String
    Dim previous As Long, scratch As Range, code As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("A1").Value2 = previous
    Range("M1:R4").Clear
    Range("N1").Formula2 = "=SEQUENCE(2,2)"
    Application.CalculateFull
    Select Case mode
        Case 0
            Range("R1").Value2 = 99
            Range("R1").ClearContents
        Case 1
            Range("R1").Interior.Color = vbRed
            Range("R1").ClearFormats
        Case 2
            Range("R1").Formula2 = "=99"
            Range("R1").Value2 = Empty
        Case 3
            Range("R1").Value2 = 99
            Range("R1").Clear
        Case 4
            Range("R1").Value2 = 99
            Range("R1").ClearContents
            Set scratch = ActiveSheet.UsedRange
        Case 5
            Range("P1").Value2 = 99
            Range("P1").ClearContents
    End Select
    CacheStage = SpillFacts(Range("N1:Q2"))
    Exit Function
Failed:
    code = Err.Number
    CacheStage = "error@stage:" & CStr(code)
End Function

Public Function CacheFinish() As String
    Dim before As String, after As String, used As String, code As Long
    On Error GoTo Failed
    before = SpillFacts(Range("N1:Q2"))
    used = ActiveSheet.UsedRange.Address
    after = SpillFacts(Range("N1:Q2"))
    CacheFinish = "before:" & before & "|used:" & used & "|after:" & after
    Application.Calculation = CLng(Range("A1").Value2)
    Range("A1").Value2 = 10
    Range("M1:R4").Clear
    Exit Function
Failed:
    code = Err.Number
    CacheFinish = "error@finish:" & CStr(code)
End Function
