' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102105
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
Public Function RefRange() As Range
 hitsC = hitsC + 1
 Set RefRange = Range("M1")
End Function
Private Function FlowSnapshot(ByVal stage As String) As String
    Dim m As Variant, n As Variant, q As Variant
    m = Range("M1").Value2: n = Range("N1").Value2: q = Range("Q1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n) & "," & CStr(VarType(q)) & ":" & CStr(q) & "]"
End Function
Public Function FlowCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, report As String, stage As String, savedError As Long
    previous = Application.Calculation
    On Error GoTo Caught
    stage = "entry"
    Application.Calculation = xlCalculationManual
    Range("M1:Q2").ClearContents
    Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Select Case kind
        Case 0: expression = "=ISREF(M1)"
        Case 1: expression = "=ISREF(INDEX(M1,1))"
        Case 2: expression = "=ISREF(OFFSET(M1,0,0))"
        Case 3: expression = "=ISREF(INDIRECT(""M1""))"
        Case 4: expression = "=ISREF(IF(TRUE,M1,O1))"
        Case 5: expression = "=ISREF(CHOOSE(1,M1,O1))"
        Case 6: expression = "=ISREF(IFS(TRUE,M1))"
        Case 7: expression = "=ISREF(SWITCH(1,1,M1))"
        Case 8: expression = "=ISREF(IFERROR(M1,0))"
        Case 9: expression = "=ISREF(IFNA(M1,0))"
        Case 10: expression = "=ISREF(LET(z,M1,z))"
        Case 11: expression = "=ISREF(TAKE(M1:M2,1))"
        Case 12: expression = "=ISREF(DROP(M1:M2,1))"
        Case 13: expression = "=ISREF(XLOOKUP(0,D1,M1,7))"
        Case 14: expression = "=ISREF(TRIMRANGE(M1:M2))"
        Case 15: expression = "=ISREF(SEQUENCE(1))"
        Case 16: expression = "=ISREF(SUM(M1))"
        Case 17: expression = "=ISREF(FlowTick(""C"",M1))"
        Case 18: expression = "=ISREF(RefRange())"
        Case 19: expression = "=ISREF(LAMBDA(r,r)(M1))"
        Case 20: expression = "=ROW(M1)"
        Case 21: expression = "=ROW(INDEX(M1,1))"
        Case 22: expression = "=ROW(OFFSET(M1,0,0))"
        Case 23: expression = "=ROW(INDIRECT(""M1""))"
        Case 24: expression = "=ROW(IF(TRUE,M1,O1))"
        Case 25: expression = "=ROW(CHOOSE(1,M1,O1))"
        Case 26: expression = "=ROW(IFS(TRUE,M1))"
        Case 27: expression = "=ROW(SWITCH(1,1,M1))"
        Case 28: expression = "=ROW(IFERROR(M1,0))"
        Case 29: expression = "=ROW(IFNA(M1,0))"
        Case 30: expression = "=ROW(LET(z,M1,z))"
        Case 31: expression = "=ROW(TAKE(M1:M2,1))"
        Case 32: expression = "=ROW(DROP(M1:M2,1))"
        Case 33: expression = "=ROW(XLOOKUP(0,D1,M1,7))"
        Case 34: expression = "=ROW(TRIMRANGE(M1:M2))"
        Case 35: expression = "=ROW(SEQUENCE(1))"
        Case 36: expression = "=ROW(SUM(M1))"
        Case 37: expression = "=ROW(FlowTick(""C"",M1))"
        Case 38: expression = "=ROW(RefRange())"
        Case 39: expression = "=ROW(LAMBDA(r,r)(M1))"
    End Select
    stage = "formula2"
    Range("N1").Formula2 = expression
    stage = "calculation"
    Range("Q1").Formula2 = "=FlowTick(""B"",N1)"
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Application.Calculate
    report = report & "|" & FlowSnapshot("again")
    Range("D1").Value2 = 1 - initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirty")
    stage = "full"
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("M1:Q2").ClearContents
    Application.Calculation = previous
    Exit Function
Caught:
    savedError = Err.Number
    FlowCase = "error@" & stage & ":" & CStr(savedError) & "|hits:" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC)
    On Error Resume Next
    Range("M1:Q2").ClearContents
    Application.Calculation = previous
End Function
