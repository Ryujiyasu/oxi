' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long, hitsS As Long
Private traceCalls As String
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102112
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
Private Function FlowSnapshot(ByVal stage As String) As String
    Dim m As Variant, n As Variant, child As Variant, other As Variant, tail As Variant, q As Variant
    m = Range("M1").Value2: n = Range("N1").Value2
    child = Range("N2").Value2: other = Range("O1").Value2
    tail = Range("N3").Value2: q = Range("Q1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsS)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n)
    FlowSnapshot = FlowSnapshot & "," & CStr(VarType(child)) & ":" & CStr(child) & "," & CStr(VarType(other)) & ":" & CStr(other)
    FlowSnapshot = FlowSnapshot & "," & CStr(VarType(tail)) & ":" & CStr(tail) & "," & CStr(VarType(q)) & ":" & CStr(q) & "]"
    FlowSnapshot = FlowSnapshot & "{" & traceCalls & "}"
    traceCalls = ""
End Function
Public Function FlowCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, child As String, report As String, stage As String, savedError As Long
    previous = Application.Calculation
    On Error GoTo Caught
    stage = "entry"
    Application.Calculation = xlCalculationManual
    Range("K1:Q4").ClearContents
    Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0: hitsS = 0
    traceCalls = ""
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Range("K1").Formula2 = "=FlowTick(""C"",17)"
    Select Case kind
        Case 0: expression = "=SEQUENCE(D1)": child = "N2"
        Case 1: expression = "=SEQUENCE(2,D1)": child = "O2"
        Case 2: expression = "=SEQUENCE(D1,2)": child = "N2"
        Case 3: expression = "=TAKE(SEQUENCE(3),D1)": child = "N2"
        Case 4: expression = "=DROP(SEQUENCE(3),D1)": child = "N2"
        Case 5: expression = "=FlowSpill(D1)": child = "N2"
        Case 6: expression = "=TRANSPOSE(FlowSpill(D1))": child = "O1"
        Case 7: expression = "=IF(D1=1,{1},{1;2})": child = "N2"
        Case 8: expression = "=CHOOSE(D1,{1},{1;2})": child = "N2"
        Case 9: expression = "=SEQUENCE(2,2)": child = "N2"
    End Select
    stage = "formula2"
    Range("N1").Formula2 = expression
    If kind = 9 Then Range("L3").Formula2 = "=SEQUENCE(1,3,11)"
    Range("Q1").Formula2 = "=FlowTick(""B""," & child & ")"
    stage = "calculation"
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Application.CalculateFull
    report = report & "|" & FlowSnapshot("stablefull")
    Range("D1").Value2 = 3 - initial
    Application.CalculateFull
    report = report & "|" & FlowSnapshot("changedfull")
    Application.CalculateFull
    report = report & "|" & FlowSnapshot("againfull")
    Range("D1").Value2 = initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirtynormal")
    stage = "full"
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("K1:Q4").ClearContents
    Application.Calculation = previous
    Exit Function
Caught:
    savedError = Err.Number
    FlowCase = "error@" & stage & ":" & CStr(savedError) & "|hits:" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsS)
    On Error Resume Next
    Range("K1:Q4").ClearContents
    Application.Calculation = previous
End Function
