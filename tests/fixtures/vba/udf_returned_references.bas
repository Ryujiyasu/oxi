' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long, hitsR As Long
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102107
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
    hitsR = hitsR + 1
    Set RefRange = Range("M1")
End Function
Public Function RefAlias(ByVal chosen As Range) As Range
    hitsR = hitsR + 1
    Set RefAlias = chosen
End Function
Public Function RefPick() As Range
    hitsR = hitsR + 1
    If Range("D2").Value2 = 0 Then
        Set RefPick = Range("M1")
    Else
        Set RefPick = Range("O1")
    End If
End Function
Private Function FlowSnapshot(ByVal stage As String) As String
    Dim m As Variant, n As Variant, q As Variant, o As Variant
    m = Range("M1").Value2: n = Range("N1").Value2
    q = Range("Q1").Value2: o = Range("O1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsR)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n)
    FlowSnapshot = FlowSnapshot & "," & CStr(VarType(q)) & ":" & CStr(q) & "," & CStr(VarType(o)) & ":" & CStr(o) & "]"
End Function
Public Function FlowCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, report As String, stage As String, savedError As Long
    previous = Application.Calculation
    On Error GoTo Caught
    stage = "entry"
    Application.Calculation = xlCalculationManual
    Range("M1:Q3").ClearContents
    Range("D1").Value2 = initial
    Range("D2").Value2 = 0
    Range("D3").Value2 = 9
    hitsA = 0: hitsB = 0: hitsC = 0: hitsR = 0
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Range("O1").Formula2 = "=FlowTick(""C"",D3)"
    Select Case kind
        Case 0: expression = "=RefRange()"
        Case 1: expression = "=RefRange()+0"
        Case 2: expression = "=SUM(RefRange())"
        Case 3: expression = "=FlowTick(""B"",RefRange())"
        Case 4: expression = "=IF(TRUE,RefRange(),0)"
        Case 5: expression = "=ROW(RefRange())"
        Case 6: expression = "=ISREF(RefRange())"
        Case 7: expression = "=RefAlias(M1)"
        Case 8: expression = "=SUM(RefAlias(M1))"
        Case 9: expression = "=RefPick()"
        Case 10: expression = "=SUM(RefPick())"
        Case 11: expression = "=ROW(RefPick())"
        Case 12: expression = "=RefAlias(IF(TRUE,M1,O1))"
        Case 13: expression = "=INDEX(RefRange(),1)"
        Case 14: expression = "=AREAS(RefRange())"
        Case 15: expression = "=IFERROR(RefRange(),0)"
    End Select
    stage = "formula2"
    Range("N1").Formula2 = expression
    Range("Q1").Formula2 = "=FlowTick(""B"",N1)"
    stage = "calculation"
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Application.Calculate
    report = report & "|" & FlowSnapshot("again")
    Range("D1").Value2 = 1 - initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirty")
    Range("D2").Value2 = 1
    Application.Calculate
    report = report & "|" & FlowSnapshot("switch")
    Range("D3").Value2 = 13
    Application.Calculate
    report = report & "|" & FlowSnapshot("other")
    stage = "full"
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("M1:Q3").ClearContents
    Application.Calculation = previous
    Exit Function
Caught:
    savedError = Err.Number
    FlowCase = "error@" & stage & ":" & CStr(savedError) & "|hits:" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsR)
    On Error Resume Next
    Range("M1:Q3").ClearContents
    Application.Calculation = previous
End Function
