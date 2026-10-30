' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102093
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
        Case 0: expression = "=FlowTick(""B"",M1)"
        Case 1: expression = "=IF(M1=0,FlowTick(""B"",5),7)"
        Case 2: expression = "=IF(M1=1,FlowTick(""B"",5),7)"
        Case 3: expression = "=IF(FlowTick(""B"",M1)=0,11,22)"
        Case 4: expression = "=IFERROR(FlowTick(""B"",M1),7)"
        Case 5: expression = "=IFERROR(1/M1,FlowTick(""B"",7))"
        Case 6: expression = "=LET(v,M1,IF(v=0,FlowTick(""B"",5),7))"
        Case 7: expression = "=LET(v,FlowTick(""B"",M1),IF(v=0,11,22))"
        Case 8: expression = "=IF(M1,FlowTick(""B"",5),FlowTick(""C"",6))"
        Case 9: expression = "=IFERROR(M1,FlowTick(""B"",7))"
        Case 10: expression = "=CHOOSE(M1+1,FlowTick(""B"",5),FlowTick(""C"",6))"
        Case 11: expression = "=IFS(M1=0,FlowTick(""B"",5),TRUE,FlowTick(""C"",6))"
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
