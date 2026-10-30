' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102094
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
        Case 1: expression = "=FlowTick(""B"",M1+0)"
        Case 2: expression = "=FlowTick(""B"",M1+1)"
        Case 3: expression = "=FlowTick(""B"",SUM(M1))"
        Case 4: expression = "=FlowTick(""B"",IF(M1=0,5,7))"
        Case 5: expression = "=FlowTick(""B"",LET(v,M1,v+0))"
        Case 6: expression = "=SUM(M1,FlowTick(""B"",5))"
        Case 7: expression = "=SUM(FlowTick(""B"",5),M1)"
        Case 8: expression = "=M1+FlowTick(""B"",5)"
        Case 9: expression = "=FlowTick(""B"",5)+M1"
        Case 10: expression = "=IF(FlowTick(""B"",M1),FlowTick(""C"",5),7)"
        Case 11: expression = "=IF(FlowTick(""B"",1),FlowTick(""C"",M1),7)"
        Case 12: expression = "=SUM(FlowTick(""B"",M1),FlowTick(""C"",5))"
        Case 13: expression = "=SUM(FlowTick(""B"",5),FlowTick(""C"",M1))"
        Case 14: expression = "=IFNA(NA()+M1,FlowTick(""B"",7))"
        Case 15: expression = "=SWITCH(M1,0,FlowTick(""B"",5),FlowTick(""C"",6))"
        Case 16: expression = "=IFS(TRUE,FlowTick(""B"",5),M1=0,FlowTick(""C"",6))"
        Case 17: expression = "=LET(unused,M1,FlowTick(""B"",5))"
        Case 18: expression = "=LET(unused,FlowTick(""B"",5),M1)"
        Case 19: expression = "=SUM(0*M1,FlowTick(""B"",5))"
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
