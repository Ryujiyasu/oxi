' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102095
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
    Dim m As Variant, n As Variant
    m = Range("M1").Value2: n = Range("N1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n) & "]"
End Function
Public Function FlowCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, report As String
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:N2").ClearContents
    Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Select Case kind
        Case 0: expression = "=FlowTick(""B"",INDEX(M1,1))"
        Case 1: expression = "=FlowTick(""B"",ROW(M1))"
        Case 2: expression = "=FlowTick(""B"",COLUMN(M1))"
        Case 3: expression = "=FlowTick(""B"",ROWS(M1))"
        Case 4: expression = "=FlowTick(""B"",ISREF(M1))"
        Case 5: expression = "=FlowTick(""B"",ISNUMBER(M1))"
        Case 6: expression = "=FlowTick(""B"",N(M1))"
        Case 7: expression = "=FlowTick(""B"",SUM(M1,0))"
        Case 8: expression = "=FlowTick(""B"",OFFSET(M1,0,0))"
        Case 9: expression = "=FlowTick(""B"",INDIRECT(""M1""))"
        Case 10: expression = "=FlowTick(""B"",XLOOKUP(0,D1,M1,7))"
        Case 11: expression = "=FlowTick(""B"",COUNTA(M1))"
        Case 12: expression = "=FlowTick(""B"",COUNT(M1))"
        Case 13: expression = "=FlowTick(""B"",LEN(M1))"
        Case 14: expression = "=FlowTick(""B"",CELL(""row"",M1))"
        Case 15: expression = "=FlowTick(""B"",IFERROR(INDEX(M1,1),7))"
        Case 16: expression = "=FlowTick(""B"",CHOOSE(1,M1))"
        Case 17: expression = "=SUM(INDEX(M1,1),FlowTick(""B"",5))"
        Case 18: expression = "=IF(INDEX(M1,1),FlowTick(""B"",5),FlowTick(""C"",6))"
        Case 19: expression = "=IF(ROW(M1)=1,FlowTick(""B"",5),7)"
    End Select
    Range("N1").Formula2 = expression
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Range("D1").Value2 = 1 - initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirty")
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("M1:N2").ClearContents
    Application.Calculation = previous
End Function
