' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102104
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
        Case 0: expression = "=ROW(1)"
        Case 1: expression = "=ROW({1})"
        Case 2: expression = "=ROW(NA())"
        Case 3: expression = "=ROW(""Sheet1"")"
        Case 4: expression = "=COLUMN(1)"
        Case 5: expression = "=COLUMN({1})"
        Case 6: expression = "=COLUMN(NA())"
        Case 7: expression = "=COLUMN(""Sheet1"")"
        Case 8: expression = "=ROWS(1)"
        Case 9: expression = "=ROWS({1})"
        Case 10: expression = "=ROWS(NA())"
        Case 11: expression = "=ROWS(""Sheet1"")"
        Case 12: expression = "=COLUMNS(1)"
        Case 13: expression = "=COLUMNS({1})"
        Case 14: expression = "=COLUMNS(NA())"
        Case 15: expression = "=COLUMNS(""Sheet1"")"
        Case 16: expression = "=ISREF(1)"
        Case 17: expression = "=ISREF({1})"
        Case 18: expression = "=ISREF(NA())"
        Case 19: expression = "=ISREF(""Sheet1"")"
        Case 20: expression = "=ISFORMULA(1)"
        Case 21: expression = "=ISFORMULA({1})"
        Case 22: expression = "=ISFORMULA(NA())"
        Case 23: expression = "=ISFORMULA(""Sheet1"")"
        Case 24: expression = "=AREAS(1)"
        Case 25: expression = "=AREAS({1})"
        Case 26: expression = "=AREAS(NA())"
        Case 27: expression = "=AREAS(""Sheet1"")"
        Case 28: expression = "=SHEET(1)"
        Case 29: expression = "=SHEET({1})"
        Case 30: expression = "=SHEET(NA())"
        Case 31: expression = "=SHEET(""Sheet1"")"
        Case 32: expression = "=SHEETS(1)"
        Case 33: expression = "=SHEETS({1})"
        Case 34: expression = "=SHEETS(NA())"
        Case 35: expression = "=SHEETS(""Sheet1"")"
        Case 36: expression = "=FORMULATEXT(1)"
        Case 37: expression = "=FORMULATEXT({1})"
        Case 38: expression = "=FORMULATEXT(NA())"
        Case 39: expression = "=FORMULATEXT(""Sheet1"")"
        Case 40: expression = "=CELL(""row"",1)"
        Case 41: expression = "=CELL(""row"",{1})"
        Case 42: expression = "=CELL(""row"",NA())"
        Case 43: expression = "=CELL(""row"",""Sheet1"")"
        Case 44: expression = "=CELL(""col"",1)"
        Case 45: expression = "=CELL(""col"",{1})"
        Case 46: expression = "=CELL(""col"",NA())"
        Case 47: expression = "=CELL(""col"",""Sheet1"")"
        Case 48: expression = "=CELL(""address"",1)"
        Case 49: expression = "=CELL(""address"",{1})"
        Case 50: expression = "=CELL(""address"",NA())"
        Case 51: expression = "=CELL(""address"",""Sheet1"")"
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
