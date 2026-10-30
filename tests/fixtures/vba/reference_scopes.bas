' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102106
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
    Dim branchIndex As Long
    previous = Application.Calculation
    On Error GoTo Caught
    stage = "entry"
    Application.Calculation = xlCalculationManual
    Range("M1:Q2").ClearContents
    Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Select Case kind
        Case 0: expression = "=ROWS(INDEX(M1:O2,,D1+1))"
        Case 1: expression = "=COLUMNS(INDEX(M1:O2,D1+1,))"
        Case 2: expression = "=ROWS(INDEX(M1:M2,,1))"
        Case 3: expression = "=ROW(INDEX(M1:M2,,1))"
        Case 4: expression = "=ISREF(LET(x_,M1,LET(x_,x_+1,x_)))"
        Case 5: expression = "=ISREF(LET(x_,M1,LET(x_,O1,x_)))"
        Case 6: expression = "=ISREF(LET(x_,M1,LET(y_,x_+1,y_)))"
        Case 7: expression = "=ISREF(LAMBDA(x_,ISREF(LET(x_,x_+1,x_)))(M1))"
        Case 8: expression = "=ROWS(INDEX(N1:N2,,1))"
        Case 9: expression = "=ISREF(LET(x_,N1,x_))"
        Case 10: expression = "=ROW(LET(x_,N1,x_))"
        Case 11: expression = "=ISREF(LET(x_,M1,x_+x_))"
        Case 12: expression = "=ROWS(INDEX(M1:O5,,M1+1))"
        Case 13: expression = "=ISREF(LAMBDA(x_,LAMBDA(x_,x_+1)(x_))(M1))"
        Case 14: expression = "=ROWS(M1:O2)"
        Case 15
            expression = "=ISREF(LET(x_0,M1"
            For branchIndex = 1 To 48
                expression = expression & ",x_" & CStr(branchIndex) & ",x_" & CStr(branchIndex-1) & "+x_" & CStr(branchIndex-1)
            Next branchIndex
            expression = expression & ",x_48))"
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
