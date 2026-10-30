' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102101
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Public Function FlowTick(ByVal label As String, ByVal chosen As Double) As Double
    Select Case label
        Case "A": hitsA = hitsA + 1
        Case "B": hitsB = hitsB + 1
        Case "C": hitsC = hitsC + 1
    End Select
    FlowTick = chosen
End Function
Private Function FlowSnapshot(ByVal stage As String) As String
    Dim m As Variant, n As Variant, q As Variant
    m = Range("M1").Value2: n = Range("N1").Value2: q = Range("Q1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n) & "," & CStr(VarType(q)) & ":" & CStr(q) & "]"
End Function
Public Function FlowCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, report As String
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:Q2").ClearContents
    Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Select Case kind
        Case 0: expression = "=FlowTick(""B"",AREAS(M1))"
        Case 1: expression = "=FlowTick(""B"",SHEET(M1))"
        Case 2: expression = "=FlowTick(""B"",SHEETS(M1))"
        Case 3: expression = "=FlowTick(""B"",LEN(FORMULATEXT(M1)))"
        Case 4: expression = "=FlowTick(""B"",ISFORMULA(M1))"
        Case 5: expression = "=AREAS(M1)"
        Case 6: expression = "=SHEET(M1)"
        Case 7: expression = "=SHEETS(M1)"
        Case 8: expression = "=LEN(FORMULATEXT(M1))"
        Case 9: expression = "=ISFORMULA(M1)"
        Case 10: expression = "=FlowTick(""B"",AREAS(N1))"
        Case 11: expression = "=FlowTick(""B"",SHEET(N1))"
        Case 12: expression = "=FlowTick(""B"",SHEETS(N1))"
        Case 13: expression = "=FlowTick(""B"",LEN(FORMULATEXT(N1)))"
        Case 14: expression = "=FlowTick(""B"",ISFORMULA(N1))"
        Case 15: expression = "=FlowTick(""B"",AREAS((M1,O1)))"
        Case 16: expression = "=FlowTick(""B"",ISREF((M1,O1)))"
        Case 17: expression = "=FlowTick(""B"",SHEET((M1,O1)))"
        Case 18: expression = "=FlowTick(""B"",SHEETS((M1,O1)))"
        Case 19: expression = "=FlowTick(""B"",LEN(FORMULATEXT((M1,O1))))"
        Case 20: expression = "=FlowTick(""B"",LEN(FORMULATEXT((O1,M1))))"
        Case 21: expression = "=FlowTick(""B"",SHEET((O1,M1)))"
        Case 22: expression = "=FlowTick(""B"",SHEETS((O1,M1)))"
        Case 23: expression = "=FlowTick(""B"",ISFORMULA((M1,O1)))"
        Case 24: expression = "=FlowTick(""B"",ISFORMULA((O1,M1)))"
        Case 25: expression = "=FlowTick(""B"",AREAS((O1,M1)))"
    End Select
    Range("N1").Formula2 = expression
    If kind >= 5 And kind <= 9 Then Range("Q1").Formula2 = "=FlowTick(""B"",N1)"
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Application.Calculate
    report = report & "|" & FlowSnapshot("again")
    Range("D1").Value2 = 1 - initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirty")
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("M1:Q2").ClearContents
    Application.Calculation = previous
End Function
