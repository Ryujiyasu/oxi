' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private evaluationCalls As Long
Private callbackLabels As String
Private lastRandomValue As Double

Private Function ScalarReport(ByVal value As Variant) As String
    Dim i As Long
    If IsError(value) Then
        Range("Z30").Value2 = value
        ScalarReport = "E:" & CStr(Application.Evaluate("ERROR.TYPE(Z30)"))
    ElseIf VarType(value) = vbString Then
        ScalarReport = "S:" & CStr(Len(value))
        For i = 1 To Len(value)
            ScalarReport = ScalarReport & ":" & CStr(AscW(Mid$(value, i, 1)))
        Next i
    Else
        ScalarReport = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function

Private Function ArrayReport(ByVal value As Variant) As String
    Dim row As Long
    Dim col As Long
    Dim columns As Long
    Dim dimensionError As Long
    If IsArray(value) Then
        On Error Resume Next
        Err.Clear
        columns = UBound(value, 2)
        dimensionError = Err.Number
        On Error GoTo 0
        If dimensionError = 9 Then
            ArrayReport = "A1:" & CStr(LBound(value)) & ":" & CStr(UBound(value))
            For row = LBound(value) To UBound(value)
                ArrayReport = ArrayReport & "|" & ScalarReport(value(row))
            Next row
        Else
            ArrayReport = "A2:" & CStr(LBound(value, 1)) & ":" & CStr(UBound(value, 1))
            ArrayReport = ArrayReport & ":" & CStr(LBound(value, 2)) & ":" & CStr(columns)
            For row = LBound(value, 1) To UBound(value, 1)
                For col = LBound(value, 2) To columns
                    ArrayReport = ArrayReport & "|" & ScalarReport(value(row, col))
                Next col
            Next row
        End If
    Else
        ArrayReport = "V|" & ScalarReport(value)
    End If
End Function

Public Function Driver() As Long
    Range("A1").Value2 = "a" & ChrW(&HD83D) & ChrW(&HDE00) & "b"
    Range("A2").Value2 = ChrW(&HD83D)
    Range("A3").Value2 = ChrW(&HDE00)
    Range("A4").Value2 = "Tokyo" & ChrW(&H6771) & ChrW(&H4EAC)
    Range("A5").Value2 = 12345
    Range("A6").Value2 = True
    Range("A7").Formula = "=1/0"
    Range("A8").ClearContents
    Range("B1").Value2 = -1
    Range("B2").Value2 = 0
    Range("B3").Value2 = 1
    Range("B4").Value2 = 2
    Range("C1").Value2 = "a"
    Range("C2").Value2 = ChrW(&HDE00)
    Range("C3").Value2 = "~?"
    Range("D1").Value2 = 1
    Range("D2").Value2 = 2
    Range("D3").Value2 = 3
    Range("E12").Value2 = 1
    Range("F12").Value2 = 2
    Range("G12").Value2 = 3
    Range("A7").Calculate
    ThisWorkbook.Names.Add Name:="OriginArray", RefersTo:="={""abc""}"
    ThisWorkbook.Names.Add Name:="OriginCell", RefersTo:="='" & ActiveSheet.Name & "'!$A$1"
    Driver = 21
End Function

Public Function BranchTrace(ByVal expression As String, ByVal mode As Long) As String
    Dim report As String, previousCalculation As Long
    previousCalculation = Application.Calculation
    evaluationCalls = 0: callbackLabels = ""
    If mode = 0 Then
        report = ArrayReport(Application.Evaluate(expression))
    ElseIf mode = 1 Then
        report = ArrayReport(ActiveSheet.Evaluate(expression))
    Else
        Application.Calculation = xlCalculationManual
        Range("M1").Formula2 = "=" & expression
        report = ArrayReport(Range("M1").Value2)
        Range("M1").ClearContents
    End If
    BranchTrace = report & "|calls:" & CStr(evaluationCalls) & "|labels:" & callbackLabels
    Application.Calculation = previousCalculation
End Function

Public Function RememberTick(ByVal selected As Double) As Double
    evaluationCalls = evaluationCalls + 1
    If evaluationCalls Mod 2 = 1 Then
        callbackLabels = callbackLabels & "F"
    ElseIf selected = lastRandomValue Then
        callbackLabels = callbackLabels & "S"
    Else
        callbackLabels = callbackLabels & "D"
    End If
    lastRandomValue = selected
    RememberTick = 1
End Function

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102085
End Function

Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function

Public Function SeedScenario(ByVal scenario As Long) As Long
    If scenario = 1 Then
        Range("D1").Value2 = 1
    ElseIf scenario = 2 Then
        Range("A7").Formula = "=1/0"
        Range("A7").Calculate
    ElseIf scenario = 3 Then
        SeedScenario = Driver()
        Exit Function
    ElseIf scenario = 4 Then
        Range("D1").Formula2 = "=RAND()"
        Range("D1").Calculate
        Range("D1").ClearContents
    ElseIf scenario = 5 Then
        ThisWorkbook.Names.Add Name:="OriginArray", RefersTo:="={""abc""}"
    ElseIf scenario = 6 Then
        Range("D1").Formula2 = "=RAND()"
    ElseIf scenario = 7 Then
        ActiveSheet.Calculate
    ElseIf scenario = 8 Then
        Range("D1").Font.Bold = True
    ElseIf scenario = 9 Then
        Range("D1:E1").Formula2 = "=RAND()"
    End If
    SeedScenario = 21
End Function
