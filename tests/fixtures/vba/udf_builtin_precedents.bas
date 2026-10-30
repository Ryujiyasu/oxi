' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102099
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
        Case 0: expression = "=D1+1"
        Case 1: expression = "=M1+1"
        Case 2: expression = "=SUM(D1)"
        Case 3: expression = "=SUM(M1)"
        Case 4: expression = "=-M1"
        Case 5: expression = "=IF(D1=0,1,2)"
        Case 6: expression = "=IF(M1=0,1,2)"
        Case 7: expression = "=INDEX(M1,1)"
        Case 8: expression = "=XLOOKUP(0,D1,M1,7)"
        Case 9: expression = "=FlowTick(""C"",M1)"
    End Select
    Range("N1").Formula2 = expression
    Range("Q1").Formula2 = "=FlowTick(""B"",N1)"
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Range("D1").Value2 = 1 - initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirty")
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("M1:Q2").ClearContents
    Application.Calculation = previous
End Function
