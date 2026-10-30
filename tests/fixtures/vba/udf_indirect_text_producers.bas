' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long, hitsT As Long
Private currentStage As String, calls As String
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102126
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Public Function TraceValue(ByVal label As String, ByVal chosen As Double) As Double
    Select Case label
        Case "A": hitsA = hitsA + 1
        Case "B": hitsB = hitsB + 1
        Case "C": hitsC = hitsC + 1
    End Select
    calls = calls & currentStage & label & ":" & CStr(chosen) & ";"
    TraceValue = chosen
End Function
Public Function TraceText(ByVal chosen As Double) As String
    hitsT = hitsT + 1
    calls = calls & currentStage & "T:" & CStr(chosen) & ";"
    If chosen = 0 Then TraceText = "M1" Else TraceText = "O1"
End Function
Private Function Snapshot(ByVal label As String) As String
    Dim m As Variant, n As Variant, o As Variant, t As Variant
    m = Range("M1").Value2: n = Range("N1").Value2
    o = Range("O1").Value2: t = Range("T1").Value2
    Snapshot = label & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC) & "," & CStr(hitsT) & ":" & calls
    Snapshot = Snapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n)
    Snapshot = Snapshot & "," & CStr(VarType(o)) & ":" & CStr(o) & "," & CStr(VarType(t)) & ":" & CStr(t) & "]"
End Function
Public Function TextCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, answer As String, code As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:T2").ClearContents
    Range("D1").Value2 = initial: Range("D2").Value2 = 0: Range("D3").Value2 = 9
    hitsA = 0: hitsB = 0: hitsC = 0: hitsT = 0
    calls = "": currentStage = "E"
    Range("M1").Formula2 = "=TraceValue(""A"",D1)"
    Range("O1").Formula2 = "=TraceValue(""C"",D3)"
    Select Case kind
        Case 0: Range("T1").Value2 = "M1"
        Case 1: Range("T1").Formula2 = "=""M""&1"
        Case 2: Range("T1").Formula2 = "=IF(D2=0,""M1"",""O1"")"
        Case 3: Range("T1").Formula2 = "=TraceText(D2)"
        Case 4: Range("T1").Value2 = "unused"
        Case 5: Range("T1").Formula2 = "=""M""&IF(D1>=0,1,1)"
    End Select
    expression = "=TraceValue(""B"",INDIRECT(T1))"
    If kind = 4 Then expression = "=TraceValue(""B"",INDIRECT(IF(D2=0,""M1"",""O1"")))"
    Range("N1").Formula2 = expression
    answer = Snapshot("entry")
    currentStage = "N": Application.Calculate
    answer = answer & "|" & Snapshot("normal")
    Range("D1").Value2 = 1 - initial
    currentStage = "D": Application.Calculate
    answer = answer & "|" & Snapshot("dirty")
    Range("D2").Value2 = 1
    currentStage = "S": Application.Calculate
    answer = answer & "|" & Snapshot("switch")
    Range("D3").Value2 = 13
    currentStage = "O": Application.Calculate
    answer = answer & "|" & Snapshot("other")
    currentStage = "F": Application.CalculateFull
    answer = answer & "|" & Snapshot("full")
    currentStage = "R": Application.CalculateFullRebuild
    TextCase = answer & "|" & Snapshot("rebuilt")
    Range("M1:T2").ClearContents
    Application.Calculation = previous
    Exit Function
Failed:
    code = Err.Number: TextCase = "error@text:" & CStr(code)
    On Error Resume Next
    Range("M1:T2").ClearContents
    Application.Calculation = previous
End Function
