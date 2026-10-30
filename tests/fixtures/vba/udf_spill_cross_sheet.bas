' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long, hitsS As Long
Private traceCalls As String
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102114
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Public Function FlowTick(ByVal label As String, ByVal chosen As Double) As Double
    If label = "B" Then traceCalls = traceCalls & "[" & CStr(VarType(chosen)) & ":" & CStr(chosen) & "]"
    Select Case label
        Case "A": hitsA = hitsA + 1
        Case "B": hitsB = hitsB + 1
        Case "C": hitsC = hitsC + 1
    End Select
    FlowTick = chosen
End Function
Public Function FlowSpill(ByVal count As Long) As Variant
    Dim values() As Variant, i As Long
    hitsS = hitsS + 1
    ReDim values(1 To count, 1 To 1)
    For i = 1 To count
        values(i, 1) = i
    Next i
    FlowSpill = values
End Function
Private Function FlowSnapshot(ByVal stage As String, ByVal a As Worksheet, ByVal b As Worksheet) As String
    Dim m As Variant, n As Variant, child As Variant, q As Variant
    m = a.Range("M1").Value2: n = a.Range("N1").Value2
    child = a.Range("N2").Value2: q = b.Range("Q1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsS)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n)
    FlowSnapshot = FlowSnapshot & "," & CStr(VarType(child)) & ":" & CStr(child) & "," & CStr(VarType(q)) & ":" & CStr(q) & "]"
    FlowSnapshot = FlowSnapshot & "{" & traceCalls & "}"
    traceCalls = ""
End Function
Public Function FlowCase(ByVal mode As Long, ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, report As String, stage As String, savedError As Long
    Dim a As Worksheet, b As Worksheet
    previous = Application.Calculation
    On Error GoTo Caught
    stage = "entry"
    Application.Calculation = xlCalculationManual
    Set a = Worksheets(1)
    Set b = Worksheets.Add
    b.Name = "Other"
    a.Range("K1:Q4").ClearContents
    b.Range("K1:Q4").ClearContents
    a.Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0: hitsS = 0
    traceCalls = ""
    a.Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    b.Range("K1").Formula2 = "=FlowTick(""C"",17)"
    If kind = 0 Then
        expression = "=SEQUENCE(D1)"
    Else
        expression = "=FlowSpill(D1)"
    End If
    stage = "formula2"
    a.Range("N1").Formula2 = expression
    b.Range("Q1").Formula2 = "=FlowTick(""B"",'" & a.Name & "'!N2)"
    stage = "calculation"
    report = FlowSnapshot("entry", a, b)
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal", a, b)
    a.Range("D1").Value2 = 3 - initial
    Select Case mode
        Case 0: a.Calculate
        Case 1: b.Calculate
        Case 2: a.Range("N1").Calculate
        Case 3: b.Range("Q1").Calculate
        Case 4: a.Calculate: b.Calculate
    End Select
    report = report & "|" & FlowSnapshot("scoped", a, b)
    Application.Calculate
    report = report & "|" & FlowSnapshot("settled", a, b)
    Application.CalculateFull
    report = report & "|" & FlowSnapshot("full", a, b)
    a.Range("D1").Value2 = initial
    b.Calculate
    report = report & "|" & FlowSnapshot("otherfirst", a, b)
    Application.Calculate
    report = report & "|" & FlowSnapshot("again", a, b)
    Application.CalculateFullRebuild
    FlowCase = report & "|" & FlowSnapshot("rebuild", a, b)
    a.Range("K1:Q4").ClearContents
    b.Range("K1:Q4").ClearContents
    Application.Calculation = previous
    Exit Function
Caught:
    savedError = Err.Number
    FlowCase = "error@" & stage & ":" & CStr(savedError) & "|hits:" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsS)
    If Not a Is Nothing Then a.Range("K1:Q4").ClearContents
    If Not b Is Nothing Then b.Range("K1:Q4").ClearContents
    Application.Calculation = previous
End Function
