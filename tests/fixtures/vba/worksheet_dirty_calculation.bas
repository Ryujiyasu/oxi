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

Public Function ByteHostArgumentCase(ByVal operation As String, ByVal slot As Long, ByVal argumentKind As Long, Optional ByVal applicationAlias As Boolean = False) As String
    Dim source As Variant
    Dim a As Variant
    Dim b As Variant
    Dim c As Variant
    Dim d As Variant
    Dim value As Variant
    Dim failure As Long
    Dim errorCodes As Variant
    errorCodes = Array(2000, 2007, 2015, 2023, 2029, 2036, 2042)
    If argumentKind = 0 Then
        Set source = Range("A1:A4")
    ElseIf argumentKind = 1 Then
        source = Array("1")
    ElseIf argumentKind = 2 Then
        Set source = Range("D1")
    ElseIf argumentKind = 4 Then
        Set source = Range("A1:A8")
    ElseIf argumentKind = 5 Then
        source = Array(CVErr(2007))
    ElseIf argumentKind = 6 Then
        Set source = Range("A7")
    ElseIf argumentKind >= 10 And argumentKind <= 16 Then
        source = Array(CVErr(errorCodes(argumentKind - 10)))
    ElseIf argumentKind >= 20 And argumentKind <= 26 Then
        source = CVErr(errorCodes(argumentKind - 20))
    ElseIf argumentKind >= 30 And argumentKind <= 36 Then
        Range("A9").Value2 = CVErr(errorCodes(argumentKind - 30))
        Set source = Range("A9")
    Else
        source = Array("1", "2")
    End If
    If operation = "ReplaceB" Then
        a = "abc": b = 2: c = 1: d = "x"
    Else
        a = "a": b = "abc": c = 1
    End If
    If IsObject(source) Then
        If slot = 0 Then Set a = source
        If slot = 1 Then Set b = source
        If slot = 2 Then Set c = source
        If slot = 3 Then Set d = source
    Else
        If slot = 0 Then a = source
        If slot = 1 Then b = source
        If slot = 2 Then c = source
        If slot = 3 Then d = source
    End If
    On Error Resume Next
    Err.Clear
    If operation = "FindB" Then
        If applicationAlias Then
            value = Application.FindB(a, b, c)
        Else
            value = Application.WorksheetFunction.FindB(a, b, c)
        End If
    ElseIf operation = "SearchB" Then
        If applicationAlias Then
            value = Application.SearchB(a, b, c)
        Else
            value = Application.WorksheetFunction.SearchB(a, b, c)
        End If
    Else
        If applicationAlias Then
            value = Application.ReplaceB(a, b, c, d)
        Else
            value = Application.WorksheetFunction.ReplaceB(a, b, c, d)
        End If
    End If
    failure = Err.Number
    On Error GoTo 0
    ByteHostArgumentCase = CStr(failure) & "|" & ArrayReport(value)
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

Public Function ByteArrayCase(ByVal expression As String) As String
    ByteArrayCase = ArrayReport(Application.Evaluate(expression))
End Function

Public Function ByteHostArrayCase(ByVal operation As String, ByVal argumentKind As Long, Optional ByVal applicationAlias As Boolean = False) As String
    Dim value As Variant
    Dim source As Variant
    Dim report As String
    On Error Resume Next
    If argumentKind = 0 Then
        Set source = Range("A1:A4")
    ElseIf argumentKind = 1 Then
        source = Range("A1:A4").Value2
    Else
        source = Array("abc", "bad", "cba")
    End If
    Err.Clear
    If operation = "FindB" Then
        If applicationAlias Then
            value = Application.FindB("a", source)
        Else
            value = Application.WorksheetFunction.FindB("a", source)
        End If
    ElseIf operation = "SearchB" Then
        If applicationAlias Then
            value = Application.SearchB("?", source)
        Else
            value = Application.WorksheetFunction.SearchB("?", source)
        End If
    Else
        If applicationAlias Then
            value = Application.ReplaceB(source, 2, 1, "x")
        Else
            value = Application.WorksheetFunction.ReplaceB(source, 2, 1, "x")
        End If
    End If
    report = CStr(Err.Number)
    On Error GoTo 0
    ByteHostArrayCase = report & "|" & ArrayReport(value)
End Function

Public Function OriginShape(ByVal expression As String) As String
    OriginShape = TypeName(Application.Evaluate(expression)) & ":" & CStr(IsObject(Application.Evaluate(expression)))
End Function

Public Function SelectionTick(ByVal selected As Variant) As Variant
    evaluationCalls = evaluationCalls + 1
    SelectionTick = selected
End Function

Public Function ReferenceTrace(ByVal expression As String, ByVal mode As Long) As String
    Dim report As String
    Dim failure As Long
    Dim previousCalculation As Long
    previousCalculation = Application.Calculation
    evaluationCalls = 0
    On Error Resume Next
    Err.Clear
    If mode = 0 Then
        report = ArrayReport(Application.Evaluate(expression))
    ElseIf mode = 1 Then
        report = ArrayReport(ActiveSheet.Evaluate(expression))
    Else
        Application.Calculation = xlCalculationManual
        Range("M1").Formula2 = "=" & expression
        Range("M1").Calculate
        report = ArrayReport(Range("M1").Value2)
        Range("M1").ClearContents
    End If
    failure = Err.Number
    ReferenceTrace = CStr(failure) & "|" & report & "|calls:" & CStr(evaluationCalls)
    On Error GoTo 0
    Application.Calculation = previousCalculation
End Function

Public Function DoubleTick(ByVal selected As Double) As Double
    evaluationCalls = evaluationCalls + 1
    DoubleTick = selected
End Function

Public Function UdfStages(ByVal expression As String) As String
    Dim previousCalculation As Long
    Dim afterWrite As Long, afterFirst As Long, afterSecond As Long
    previousCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1").ClearContents
    evaluationCalls = 0
    Range("M1").Formula2 = "=" & expression
    afterWrite = evaluationCalls
    Range("M1").Calculate
    afterFirst = evaluationCalls
    Range("M1").Calculate
    afterSecond = evaluationCalls
    UdfStages = CStr(afterWrite) & ":" & CStr(afterFirst) & ":" & CStr(afterSecond) & "|" & ArrayReport(Range("M1").Value2)
    Range("M1").ClearContents
    Application.Calculation = previousCalculation
End Function

Public Function CounterTick() As Double
    evaluationCalls = evaluationCalls + 1
    CounterTick = evaluationCalls
End Function

Public Function ValueStages(ByVal expression As String, ByVal useFormula2 As Boolean) As String
    Dim previousCalculation As Long
    Dim beforeRead As Long
    Dim report As String, stages As String
    previousCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1").ClearContents
    evaluationCalls = 0
    If useFormula2 Then
        Range("M1").Formula2 = "=" & expression
    Else
        Range("M1").Formula = "=" & expression
    End If
    beforeRead = evaluationCalls
    report = ArrayReport(Range("M1").Value2)
    stages = CStr(beforeRead) & "~" & CStr(evaluationCalls) & ":" & report
    Range("M1").Calculate
    beforeRead = evaluationCalls
    report = ArrayReport(Range("M1").Value2)
    stages = stages & "|" & CStr(beforeRead) & "~" & CStr(evaluationCalls) & ":" & report
    Range("M1").Calculate
    beforeRead = evaluationCalls
    report = ArrayReport(Range("M1").Value2)
    ValueStages = stages & "|" & CStr(beforeRead) & "~" & CStr(evaluationCalls) & ":" & report
    Range("M1").ClearContents
    Application.Calculation = previousCalculation
End Function

Private Function ShapeAndValue(ByVal value As Variant) As String
    Dim report As String
    report = TypeName(value) & ":" & CStr(IsObject(value))
    If IsObject(value) Then
        report = report & "|" & ArrayReport(value.Value2)
    Else
        report = report & "|" & ArrayReport(value)
    End If
    ShapeAndValue = report
End Function

Public Function RootProbe(ByVal expression As String, ByVal worksheetEntry As Boolean) As String
    Dim report As String
    Dim failure As Long
    evaluationCalls = 0
    On Error Resume Next
    Err.Clear
    If worksheetEntry Then
        report = ShapeAndValue(ActiveSheet.Evaluate(expression))
    Else
        report = ShapeAndValue(Application.Evaluate(expression))
    End If
    failure = Err.Number
    RootProbe = CStr(failure) & "|" & report & "|calls:" & CStr(evaluationCalls)
    On Error GoTo 0
End Function

Public Function TaggedTick(ByVal label As String, ByVal selected As Double) As Double
    evaluationCalls = evaluationCalls + 1
    callbackLabels = callbackLabels & label
    TaggedTick = selected
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

Public Function CellScopeTrace() As String
    Dim previousCalculation As Long, report As String
    previousCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    evaluationCalls = 0
    Range("M1").Formula2 = "=CounterTick()"
    report = CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Range("N1").Formula2 = "=CounterTick()"
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("M1").Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("N1").Calculate
    CellScopeTrace = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("M1:N3").ClearContents
    Application.Calculation = previousCalculation
End Function

Public Function TaggedError(ByVal label As String) As Variant
    evaluationCalls = evaluationCalls + 1
    callbackLabels = callbackLabels & label
    TaggedError = CVErr(2015)
End Function

Public Function AssignmentScope(ByVal automatic As Boolean) As String
    Dim previousCalculation As Long, savedInput As Variant, report As String
    previousCalculation = Application.Calculation
    savedInput = Range("D1").Value2
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    Range("M1").Formula2 = "=DoubleTick(D1)"
    Range("N1").Formula2 = "=DoubleTick(D2)"
    If automatic Then Application.Calculation = xlCalculationAutomatic
    evaluationCalls = 0
    Range("D1").Value2 = 7
    report = CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("M1").Calculate
    AssignmentScope = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("M1:N3").ClearContents
    Range("D1").Value2 = savedInput
    Application.Calculation = previousCalculation
End Function

Public Function ReferenceStages(ByVal expression As String, ByVal automatic As Boolean) As String
    Dim previousCalculation As Long, savedInput As Variant, report As String
    previousCalculation = Application.Calculation
    savedInput = Range("D1").Value2
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    If automatic Then Application.Calculation = xlCalculationAutomatic
    evaluationCalls = 0
    Range("M1").Formula2 = "=" & expression
    report = CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Range("M1").Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Range("M1").Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Range("D1").Value2 = 7
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Range("M1").Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    ActiveSheet.Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Application.Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Application.CalculateFull
    ReferenceStages = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2)
    Range("M1:N3").ClearContents
    Range("D1").Value2 = savedInput
    Application.Calculation = previousCalculation
End Function

Public Function VolatileCounter() As Double
    Application.Volatile
    evaluationCalls = evaluationCalls + 1
    VolatileCounter = evaluationCalls
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

Public Function NestedTick(ByVal selected As Double) As Double
    evaluationCalls = evaluationCalls + 1
    callbackLabels = callbackLabels & "O"
    NestedTick = selected + ActiveSheet.Evaluate("TaggedTick(""A"",1)")
End Function

Public Function CalculationSteps(ByVal automatic As Boolean, ByVal operations As String, ByVal expression As String) As String
    Dim previousCalculation As Long, savedInput As Variant, report As String, i As Long
    previousCalculation = Application.Calculation
    savedInput = Range("D1").Value2
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    If automatic Then Application.Calculation = xlCalculationAutomatic
    evaluationCalls = 0
    Range("M1").Formula2 = "=" & expression
    report = CStr(evaluationCalls) & ":" & ArrayReport(Range("M1").Value2)
    For i = 1 To Len(operations)
        Select Case Mid$(operations, i, 1)
        Case "R": Range("M1").Calculate
        Case "S": ActiveSheet.Calculate
        Case "A": Application.Calculate
        Case "F": Application.CalculateFull
        Case "D": Range("M1").Dirty
        Case "W": Range("D1").Value2 = 7
        End Select
        report = report & "|" & CStr(evaluationCalls) & ":" & ArrayReport(Range("M1").Value2)
    Next i
    CalculationSteps = report
    Range("M1:N3").ClearContents
    Range("D1").Value2 = savedInput
    Application.Calculation = previousCalculation
End Function

Public Function DirtyIsolation() As String
    Dim previousCalculation As Long, input1 As Variant, input2 As Variant, report As String
    previousCalculation = Application.Calculation
    input1 = Range("D1").Value2: input2 = Range("D2").Value2
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    evaluationCalls = 0
    Range("M1").Formula2 = "=DoubleTick(D1)"
    Range("N1").Formula2 = "=DoubleTick(D2)"
    Application.Calculate
    report = CStr(evaluationCalls)
    Range("D1").Value2 = 7: Range("D2").Value2 = 8
    Range("M1").Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Application.Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Application.Calculate
    DirtyIsolation = report & "|" & CStr(evaluationCalls)
    Range("M1:N3").ClearContents
    Range("D1").Value2 = input1: Range("D2").Value2 = input2
    Application.Calculation = previousCalculation
End Function

Public Function UdfArgumentKind(ByVal value As Variant) As Double
    evaluationCalls = evaluationCalls + 1
    If IsObject(value) Then
        UdfArgumentKind = 1
    ElseIf IsArray(value) Then
        UdfArgumentKind = 64
    Else
        UdfArgumentKind = VarType(value)
    End If
End Function

Public Function DynamicVolatile(ByVal enabled As Boolean, ByVal label As String) As Double
    Application.Volatile enabled
    evaluationCalls = evaluationCalls + 1
    callbackLabels = callbackLabels & label
    DynamicVolatile = evaluationCalls
End Function

Public Function VolatileCells(ByVal first As Boolean, ByVal second As Boolean, ByVal automatic As Boolean) As String
    Dim previousCalculation As Long, report As String
    previousCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    If automatic Then Application.Calculation = xlCalculationAutomatic
    evaluationCalls = 0: callbackLabels = ""
    Range("M1").Formula2 = "=DynamicVolatile(" & CStr(first) & ",""A"")"
    Range("N1").Formula2 = "=DynamicVolatile(" & CStr(second) & ",""B"")"
    Application.Calculate
    evaluationCalls = 0: callbackLabels = ""
    Application.Calculate
    report = CStr(evaluationCalls) & ":" & callbackLabels
    Application.Calculate
    VolatileCells = report & "|" & CStr(evaluationCalls) & ":" & callbackLabels
    Range("M1:N3").ClearContents
    Application.Calculation = previousCalculation
End Function

Public Function NamedDirtyDependency() As String
    Dim previousCalculation As Long, savedSourceValue As Variant, report As String
    previousCalculation = Application.Calculation
    savedSourceValue = Range("D1").Value2
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    ThisWorkbook.Names.Add Name:="DirtyInput", RefersTo:="='" & ActiveSheet.Name & "'!$D$1"
    evaluationCalls = 0
    Range("M1").Formula2 = "=DoubleTick(DirtyInput)"
    Range("N1").Formula2 = "=DoubleTick(M1)"
    Application.Calculate
    report = CStr(evaluationCalls)
    Range("D1").Value2 = 7
    Range("M1").Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Application.Calculate
    NamedDirtyDependency = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("M1:N3").ClearContents
    ThisWorkbook.Names("DirtyInput").Delete
    Range("D1").Value2 = savedSourceValue
    Application.Calculation = previousCalculation
End Function

Public Function TableDirtyDependency() As String
    Dim previousCalculation As Long, report As String, table As ListObject
    previousCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("H1:H3").ClearContents: Range("M1:N3").ClearContents
    Range("H1").Value2 = "Value"
    Range("H2").Value2 = 1: Range("H3").Value2 = 2
    Set table = ActiveSheet.ListObjects.Add(xlSrcRange, Range("H1:H3"), , xlYes)
    table.Name = "DirtyValues"
    evaluationCalls = 0
    Range("M1").Formula2 = "=DoubleTick(SUM(DirtyValues[Value]))"
    Range("N1").Formula2 = "=DoubleTick(M1)"
    Application.Calculate
    report = CStr(evaluationCalls)
    Range("H2").Value2 = 7
    Application.Calculate
    report = report & "|" & CStr(evaluationCalls) & ":" & CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Application.Calculate
    TableDirtyDependency = report & "|" & CStr(evaluationCalls)
    Range("M1:N3").ClearContents
    table.Delete
    Application.Calculation = previousCalculation
End Function

Public Function BeginDirtyState() As String
    Application.Calculation = xlCalculationManual
    Range("M1:N3").ClearContents
    Range("D1").Value2 = 1: Range("D2").Value2 = 2
    Range("M1").Formula2 = "=DoubleTick(D1)"
    Range("N1").Formula2 = "=DoubleTick(D2)"
    Application.Calculate
    Range("D1").Value2 = 7: Range("D2").Value2 = 8
    Range("M1").Calculate
    BeginDirtyState = CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
End Function

Public Function FinishDirtyState() As String
    Application.Calculate
    FinishDirtyState = CStr(Range("M1").Value2) & ":" & CStr(Range("N1").Value2)
    Range("M1:N3").ClearContents
    Range("D1").Value2 = 1: Range("D2").Value2 = 2
    Application.Calculation = xlCalculationAutomatic
End Function
