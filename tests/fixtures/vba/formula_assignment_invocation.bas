' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102145
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Private Function TypedValue(ByVal value As Variant) As String
    If IsEmpty(value) Then
        TypedValue = "0:Empty"
    ElseIf IsNull(value) Then
        TypedValue = "1:Null"
    Else
        TypedValue = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function
Private Function CellState(ByVal cell As Range) As String
    CellState = "value=" & TypedValue(cell.Value2) & ";formula=" & TypedValue(cell.Formula2)
    CellState = CellState & ";r1c1=" & TypedValue(cell.Formula2R1C1) & ";has=" & CStr(cell.HasFormula)
End Function
Private Function PairState() As String
    PairState = "N{" & CellState(Range("N1")) & "}O{" & CellState(Range("O1")) & "}"
End Function
Private Function ArgumentText(ByVal kind As Long, ByVal route As Long) As String
    Dim d As String, e As String, f As String, band As String
    d = "D1": e = "E1": f = "F1"
    If route Mod 2 = 1 Then d = "R1C4": e = "R1C5": f = "R1C6"
    band = d & ":" & e
    Select Case kind
        Case 0: ArgumentText = d
        Case 1: ArgumentText = """tokyo"""
        Case 2: ArgumentText = "1"
        Case 3: ArgumentText = "TRUE"
        Case 4: ArgumentText = "{1,2}"
        Case 5: ArgumentText = "#N/A"
        Case 6: ArgumentText = "NA()"
        Case 7: ArgumentText = "1/0"
        Case 8: ArgumentText = d & "&"""""
        Case 9: ArgumentText = "IF(TRUE," & d & "," & e & ")"
        Case 10: ArgumentText = "IF(TRUE," & d & ",""x"")"
        Case 11: ArgumentText = "IF(FALSE," & d & ",""x"")"
        Case 12: ArgumentText = "CHOOSE(1," & d & "," & e & ")"
        Case 13: ArgumentText = "INDEX(" & band & ",1,1)"
        Case 14: ArgumentText = "INDEX({1,2},1,1)"
        Case 15: ArgumentText = "TAKE(" & band & ",1)"
        Case 16: ArgumentText = "FILTER(" & band & ",{TRUE,TRUE})"
        Case 17: ArgumentText = "(" & d & "," & e & ")"
        Case 18: ArgumentText = f
        Case 19: ArgumentText = "INDIRECT(""D1"")"
        Case 20: ArgumentText = "LET(held," & d & ",held)"
        Case 21: ArgumentText = "IF(TRUE,""x"",""y"")"
        Case 22: ArgumentText = "OFFSET(" & d & ",0,0)"
        Case 23: ArgumentText = "CHOOSE(2," & d & ",""x"")"
        Case 24: ArgumentText = "CHOOSE(1," & d & ",""x"")"
        Case 25: ArgumentText = "TAKE({1,2},1)"
        Case 26: ArgumentText = """D1"""
        Case 27: ArgumentText = "INDIRECT(""missing!"")"
        Case 28: ArgumentText = d & "#"
        Case 29: ArgumentText = "LET(held,""x"",held)"
        Case 30: ArgumentText = "LAMBDA(held,held)(" & d & ")"
        Case 31: ArgumentText = "IFERROR(" & d & "," & e & ")"
    End Select
End Function
Public Function AssignmentCase(ByVal kind As Long, ByVal route As Long) As String
    Dim previous As Long, expression As String, member As String, answer As String, fault As Long, calculateFault As Long
    Dim target As Range, ignored As Variant
    On Error GoTo Failed
    previous = Application.Calculation: Application.Calculation = xlCalculationManual
    Range("D1").Value2 = "tokyo": Range("E1").Value2 = 12: Range("F1").Value2 = CVErr(2042)
    Set target = Range("N1"): target.Value2 = 401
    member = "Formula2": If route Mod 2 = 1 Then member = "Formula2R1C1"
    expression = "=PHONETIC(" & ArgumentText(kind, route) & ")"
    answer = "expression=" & expression & "|before{" & CellState(target) & "}"
    On Error Resume Next
    Err.Clear: ignored = CallByName(target, member, VbLet, expression): fault = Err.Number: Err.Clear
    On Error GoTo Failed
    answer = answer & "|fault=" & CStr(fault) & "|immediate{" & CellState(target) & "}"
    On Error Resume Next
    Err.Clear: Application.Calculate: calculateFault = Err.Number: Err.Clear
    On Error GoTo Failed
    answer = answer & "|calculateFault=" & CStr(calculateFault) & "|calculated{" & CellState(target) & "}"
    Application.Calculation = previous: AssignmentCase = answer
    Exit Function
Failed:
    fault = Err.Number: AssignmentCase = "construction-fault=" & CStr(fault)
    On Error Resume Next
    Application.Calculation = previous
End Function
Public Function AtomicCase(ByVal kind As Long, ByVal route As Long) As String
    Dim previous As Long, member As String, answer As String, fault As Long, calculateFault As Long
    Dim target As Range, expression As Variant, ignored As Variant, matrix(1 To 1, 1 To 2) As Variant
    On Error GoTo Failed
    previous = Application.Calculation: Application.Calculation = xlCalculationManual
    Range("D1").Value2 = "tokyo": Range("E1").Value2 = 12
    Range("N1").Value2 = 401: Range("O1").Value2 = 402
    member = "Formula2": If route Mod 2 = 1 Then member = "Formula2R1C1"
    Set target = Range("N1:O1")
    Select Case kind
        Case 0
            Set target = Range("N1"): target.Formula2 = "=17+4": expression = "=PHONETIC(1)"
        Case 1
            expression = "=PHONETIC(1)"
        Case 2
            matrix(1, 1) = "=1+2": matrix(1, 2) = "=PHONETIC(1)": expression = matrix
        Case 3
            matrix(1, 1) = "=PHONETIC(1)": matrix(1, 2) = "=3+4": expression = matrix
        Case 4
            matrix(1, 1) = "=1+2": matrix(1, 2) = "=3+4": expression = matrix
        Case 5
            matrix(1, 1) = "=PHONETIC(D1)": If route Mod 2 = 1 Then matrix(1, 1) = "=PHONETIC(R1C4)"
            matrix(1, 2) = "=PHONETIC(1)": expression = matrix
    End Select
    answer = "before{" & PairState() & "}"
    On Error Resume Next
    Err.Clear
    If route < 2 Then
        ignored = CallByName(target, member, VbLet, expression)
    ElseIf route = 2 Then
        target.Formula2 = expression
    Else
        target.Formula2R1C1 = expression
    End If
    fault = Err.Number: Err.Clear
    On Error GoTo Failed
    answer = answer & "|fault=" & CStr(fault) & "|immediate{" & PairState() & "}"
    On Error Resume Next
    Err.Clear: Application.Calculate: calculateFault = Err.Number: Err.Clear
    On Error GoTo Failed
    answer = answer & "|calculateFault=" & CStr(calculateFault) & "|calculated{" & PairState() & "}"
    Application.Calculation = previous: AtomicCase = answer
    Exit Function
Failed:
    fault = Err.Number: AtomicCase = "construction-fault=" & CStr(fault)
    On Error Resume Next
    Application.Calculation = previous
End Function
