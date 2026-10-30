' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit

Private Function UnitText(ByVal text As String) As String

    Dim i As Long

    UnitText = CStr(Len(text))

    For i = 1 To Len(text)

        UnitText = UnitText & ":" & CStr(AscW(Mid$(text, i, 1)))

    Next i

End Function

Public Function SetupCells() As String

    Dim values As Variant

    Dim matrix(1 To 12, 1 To 1) As Variant

    Dim i As Long

    values = Array(ChrW(&HD83D), ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00), ChrW(&HD83D) & ChrW(0) & ChrW(&HDE00), ChrW(0), "_xD83D_", "_x005F_", "_x0000_", "A" & vbCr & "B", "A" & vbLf & "B", "<&>", "plain")

    For i = 1 To 12

        Cells(i, 1).value = values(i - 1)

        Cells(i, 2).Value2 = values(i - 1)

        matrix(i, 1) = values(i - 1)

    Next i

    Range("C1:C12").Value2 = matrix

    Range("A1:A12").Copy Destination:=Range("D1")

    Range("C1:C12").Copy

    Range("E1").PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    SetupCells = ReadCells()

End Function

Public Function ReadCells() As String

    Dim row As Long

    Dim col As Long

    Dim report As String

    For row = 1 To 12

        If row > 1 Then report = report & "|"

        For col = 1 To 5

            If col > 1 Then report = report & "/"

            report = report & UnitText(Cells(row, col).Value2)

        Next col

    Next row

    ReadCells = report

End Function

Private Function ReportValue(ByVal value As Variant) As String

    If IsError(value) Then

        ReportValue = "E"

    ElseIf VarType(value) = vbString Then

        ReportValue = "S:" & UnitText(value)

    Else

        ReportValue = CStr(VarType(value)) & ":" & CStr(value)

    End If

End Function

Public Function FormulaCases() As String

    Dim expressions As Variant

    Dim expression As Variant

    Dim report As String

    expressions = Array("=A1", "=A1&""""", "=LEN(A1)", "=LEN(A3)", "=LEFT(A3,1)", "=RIGHT(A3,1)", "=MID(A3,2,1)", "=ISTEXT(A1)", "=TYPE(A1)", "=A1=A1", "=A1=A2", "=A1<A2", "=A1>""a""", "=UNICODE(A1)", "=UNICODE(A3)", "=EXACT(A1,A2)", "=SUBSTITUTE(A3,A1,""x"")", "=SEARCH(A2,A3)", "=FIND(A2,A3)", "=UPPER(A1)", "=TRIM(A1)", "=T(A1)")

    For Each expression In expressions

        If Len(report) > 0 Then report = report & "|"

        report = report & ReportValue(Application.Evaluate(expression))

    Next expression

    FormulaCases = report

End Function

Public Function CharacterCases() As String

    Dim report As String

    report = CStr(Range("A1").Characters.Count) & ":" & CStr(Range("A3").Characters.Count)

    report = report & "|" & UnitText(Range("A1").text)

    report = report & "|" & UnitText(Range("A3").Characters(1, 1).text)

    report = report & "|" & UnitText(Range("A3").Characters(2, 1).text)

    Range("A3").Characters(1, 1).text = "x"

    report = report & "|" & UnitText(Range("A3").Value2)

    Range("A1").Characters(1, 1).Font.Bold = True

    report = report & "|" & UnitText(Range("A1").Value2)

    CharacterCases = report

End Function

Public Function Modes() As String

    Dim pair As String

    Dim value As Variant

    Dim report As String

    pair = Range("A3").Value2

    report = CStr(Len(Range("A1").Value2)) & ":" & CStr(Len(pair))

    On Error Resume Next

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.Len(Range("A1").Value2)

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.Len(pair)

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.Left(pair, 1)

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.Right(pair, 1)

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.Mid(pair, 2, 1)

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    On Error GoTo 0

    Range("G1").Formula = "=LEN(A1)"

    Range("G2").Formula = "=LEN(A3)"

    Range("G3").Formula2 = "=LEN(A1)"

    Range("G4").Formula2 = "=LEN(A3)"

    Range("G1:G4").Calculate

    report = report & "|" & CStr(Range("G1").Value2) & ":" & CStr(Range("G2").Value2) & ":" & CStr(Range("G3").Value2) & ":" & CStr(Range("G4").Value2)

    Modes = report

End Function

Public Function AfterEdit() As String

    AfterEdit = CStr(Range("G1").Value2) & ":" & CStr(Range("G2").Value2) & ":" & CStr(Range("G3").Value2) & ":" & CStr(Range("G4").Value2) & "|" & CStr(Range("A1").Font.Bold)

End Function

Public Function SupportedFunctions() As String

    Dim value As Variant

    Dim report As String

    On Error Resume Next

    Err.Clear

    value = Application.WorksheetFunction.Find(ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00))

    report = CStr(Err.Number) & ":" & ReportValue(value)

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.Search(ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00))

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    Err.Clear

    value = Empty

    value = Application.WorksheetFunction.IsText(ChrW(&HD83D))

    report = report & "|" & CStr(Err.Number) & ":" & ReportValue(value)

    On Error GoTo 0

    SupportedFunctions = report

End Function

Public Function TextExtensions() As String

    Dim samples As Variant

    Dim sample As Variant

    Dim expressions As Variant

    Dim expression As Variant

    Dim report As String

    samples = Array("", ChrW(&HD83D), ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00), "a" & ChrW(&HD83D) & "b", " " & ChrW(&HD83D) & " ", "a" & ChrW(&HD83D) & ChrW(&HDE00) & "b", ChrW(&HD83D) & ChrW(&HDE00) & ChrW(&HD83D), "A" & ChrW(1) & "b", "a_B" & ChrW(&HDE00) & "9", "Mixed Word")

    expressions = Array("=CONCAT(J20,""x"")", "=CONCATENATE(J20,""x"")", "=CONCAT(J20,J20)", "=TEXTJOIN(""|"",FALSE,J20,J20)", "=TEXTJOIN("""",TRUE,J20,J20)", "=REPT(J20,2)", "=REPT(J20,0)", "=CLEAN(J20)", "=PROPER(J20)", "=UNICODE(J20)", "=LENB(J20)", "=LEFTB(J20,1)", "=LEFTB(J20,2)", "=RIGHTB(J20,1)", "=MIDB(J20,1,1)", "=TEXTBEFORE(J20,K20,1,0,0,""none"")", "=TEXTAFTER(J20,K20,1,0,0,""none"")", "=VALUETOTEXT(J20)", "=TEXT(J20,""@"")", "=T(J20)")

    Range("K20").Value2 = ChrW(&HD83D)

    For Each sample In samples

        If Len(report) > 0 Then report = report & "#"

        Range("J20").Value2 = sample

        report = report & UnitText(sample) & "="

        For Each expression In expressions

            report = report & ReportValue(Application.Evaluate(expression)) & "|"

        Next expression

    Next sample

    TextExtensions = report

End Function

Public Function ByteCases() As String

    Dim samples As Variant

    Dim sample As Variant

    Dim expressions As Variant

    Dim expression As Variant

    Dim report As String

    samples = Array(ChrW(&HE9), ChrW(&H3C0), ChrW(&H6F22), ChrW(&H3A9), ChrW(&HFF21), ChrW(&HFF71), ChrW(&HA5), ChrW(&H20AC), ChrW(&HD83D) & ChrW(&HDE00), ChrW(&HD842) & ChrW(&HDFB7), ChrW(&HD83D), ChrW(&HDE00), ChrW(&H3042) & ChrW(&HD83D), "a" & ChrW(&H3042) & ChrW(&HD83D) & ChrW(&HDE00) & "b")

    expressions = Array("=LENB(J20)", "=LEFTB(J20,1)", "=RIGHTB(J20,1)", "=MIDB(J20,1,1)", "=LEFTB(J20,2)", "=RIGHTB(J20,2)", "=MIDB(J20,2,1)", "=MIDB(J20,2,2)")

    For Each sample In samples

        If Len(report) > 0 Then report = report & "#"

        Range("J20").Value2 = sample

        report = report & UnitText(sample) & "="

        For Each expression In expressions

            report = report & ReportValue(Application.Evaluate(expression)) & "|"

        Next expression

    Next sample

    ByteCases = report

End Function

Public Function ByteBoundaryControls() As String
    Dim samples As Variant
    Dim sample As Variant
    Dim expressions As Variant
    Dim expression As Variant
    Dim value As Variant
    Dim report As String
    samples = Array("", ChrW(&HD83D), ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00) & ChrW(&HD83D), ChrW(&HD83D) & ChrW(&HD83D) & ChrW(&HDE00), ChrW(&HDE00) & ChrW(&HD83D) & ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00) & ChrW(&HD83D) & ChrW(&HDE00) & ChrW(&HD83D), "a" & ChrW(&HD83D) & ChrW(&HDE00) & "b", "a" & ChrW(&HD83D), ChrW(&HD83D) & ChrW(&HD83D), ChrW(&HDE00) & ChrW(&HDE00), "~x", "~")
    expressions = Array("LEFTB(J20,1)", "LEFTB(J20,2)", "LEFTB(J20,3)", "LEFTB(J20,4)", "LEFTB(J20,5)", "RIGHTB(J20,1)", "RIGHTB(J20,2)", "RIGHTB(J20,3)", "RIGHTB(J20,4)", "RIGHTB(J20,5)", "LEFT(J20,1)", "LEFT(J20,2)", "LEFT(J20,3)", "LEFT(J20,4)", "LEFT(J20,5)", "RIGHT(J20,1)", _
        "RIGHT(J20,2)", "RIGHT(J20,3)", "RIGHT(J20,4)", "RIGHT(J20,5)", "MIDB(J20,1,1)", "MIDB(J20,1,2147483647)", "MIDB(J20,1,2147483647.9)", "MIDB(J20,1,2147483648)", "MIDB(J20,1,4294967295)", "MIDB(J20,2147483647,1)", "MIDB(J20,2147483647,2147483647)", "MIDB(J20,2147483647,2147483647.9)", "MIDB(J20,2147483647,2147483648)", "MIDB(J20,2147483647,4294967295)", "MIDB(J20,2147483647.9,1)", "MIDB(J20,2147483647.9,2147483647)", _
        "MIDB(J20,2147483647.9,2147483647.9)", "MIDB(J20,2147483647.9,2147483648)", "MIDB(J20,2147483647.9,4294967295)", "MIDB(J20,2147483648,1)", "MIDB(J20,2147483648,2147483647)", "MIDB(J20,2147483648,2147483647.9)", "MIDB(J20,2147483648,2147483648)", "MIDB(J20,2147483648,4294967295)", "MIDB(J20,4294967295,1)", "MIDB(J20,4294967295,2147483647)", "MIDB(J20,4294967295,2147483647.9)", "MIDB(J20,4294967295,2147483648)", "MIDB(J20,4294967295,4294967295)", "MIDB(J20,4294967296,1)", "MIDB(J20,4294967296,2147483647)", "MIDB(J20,4294967296,2147483647.9)", _
        "MIDB(J20,4294967296,2147483648)", "MIDB(J20,4294967296,4294967295)", "REPLACEB(J20,1,1,K20)", "REPLACEB(J20,1,2147483647,K20)", "REPLACEB(J20,1,2147483647.9,K20)", "REPLACEB(J20,1,2147483648,K20)", "REPLACEB(J20,1,4294967295,K20)", "REPLACEB(J20,2147483647,1,K20)", "REPLACEB(J20,2147483647,2147483647,K20)", "REPLACEB(J20,2147483647,2147483647.9,K20)", "REPLACEB(J20,2147483647,2147483648,K20)", "REPLACEB(J20,2147483647,4294967295,K20)", "REPLACEB(J20,2147483647.9,1,K20)", "REPLACEB(J20,2147483647.9,2147483647,K20)", "REPLACEB(J20,2147483647.9,2147483647.9,K20)", "REPLACEB(J20,2147483647.9,2147483648,K20)", _
        "REPLACEB(J20,2147483647.9,4294967295,K20)", "REPLACEB(J20,2147483648,1,K20)", "REPLACEB(J20,2147483648,2147483647,K20)", "REPLACEB(J20,2147483648,2147483647.9,K20)", "REPLACEB(J20,2147483648,2147483648,K20)", "REPLACEB(J20,2147483648,4294967295,K20)", "REPLACEB(J20,4294967295,1,K20)", "REPLACEB(J20,4294967295,2147483647,K20)", "REPLACEB(J20,4294967295,2147483647.9,K20)", "REPLACEB(J20,4294967295,2147483648,K20)", "REPLACEB(J20,4294967295,4294967295,K20)", "REPLACEB(J20,4294967296,1,K20)", "REPLACEB(J20,4294967296,2147483647,K20)", "REPLACEB(J20,4294967296,2147483647.9,K20)", "REPLACEB(J20,4294967296,2147483648,K20)", "REPLACEB(J20,4294967296,4294967295,K20)", _
        "FINDB("""",J20,1)", "FINDB("""",J20,2)", "FINDB("""",J20,3)", "FINDB("""",J20,4)", "FINDB("""",J20,5)", "FINDB(K20,J20,1)", "FINDB(K20,J20,2)", "FINDB(K20,J20,3)", "FINDB(K20,J20,4)", "FINDB(K20,J20,5)", "FINDB(L20,J20,1)", "FINDB(L20,J20,2)", "FINDB(L20,J20,3)", "FINDB(L20,J20,4)", "FINDB(L20,J20,5)", "FINDB(""?"",J20,1)", _
        "FINDB(""?"",J20,2)", "FINDB(""?"",J20,3)", "FINDB(""?"",J20,4)", "FINDB(""?"",J20,5)", "FINDB(""*"",J20,1)", "FINDB(""*"",J20,2)", "FINDB(""*"",J20,3)", "FINDB(""*"",J20,4)", "FINDB(""*"",J20,5)", "FINDB(""~"",J20,1)", "FINDB(""~"",J20,2)", "FINDB(""~"",J20,3)", "FINDB(""~"",J20,4)", "FINDB(""~"",J20,5)", "FINDB(""~x"",J20,1)", "FINDB(""~x"",J20,2)", _
        "FINDB(""~x"",J20,3)", "FINDB(""~x"",J20,4)", "FINDB(""~x"",J20,5)", "FINDB(""~~"",J20,1)", "FINDB(""~~"",J20,2)", "FINDB(""~~"",J20,3)", "FINDB(""~~"",J20,4)", "FINDB(""~~"",J20,5)", "SEARCHB("""",J20,1)", "SEARCHB("""",J20,2)", "SEARCHB("""",J20,3)", "SEARCHB("""",J20,4)", "SEARCHB("""",J20,5)", "SEARCHB(K20,J20,1)", "SEARCHB(K20,J20,2)", "SEARCHB(K20,J20,3)", _
        "SEARCHB(K20,J20,4)", "SEARCHB(K20,J20,5)", "SEARCHB(L20,J20,1)", "SEARCHB(L20,J20,2)", "SEARCHB(L20,J20,3)", "SEARCHB(L20,J20,4)", "SEARCHB(L20,J20,5)", "SEARCHB(""?"",J20,1)", "SEARCHB(""?"",J20,2)", "SEARCHB(""?"",J20,3)", "SEARCHB(""?"",J20,4)", "SEARCHB(""?"",J20,5)", "SEARCHB(""*"",J20,1)", "SEARCHB(""*"",J20,2)", "SEARCHB(""*"",J20,3)", "SEARCHB(""*"",J20,4)", _
        "SEARCHB(""*"",J20,5)", "SEARCHB(""~"",J20,1)", "SEARCHB(""~"",J20,2)", "SEARCHB(""~"",J20,3)", "SEARCHB(""~"",J20,4)", "SEARCHB(""~"",J20,5)", "SEARCHB(""~x"",J20,1)", "SEARCHB(""~x"",J20,2)", "SEARCHB(""~x"",J20,3)", "SEARCHB(""~x"",J20,4)", "SEARCHB(""~x"",J20,5)", "SEARCHB(""~~"",J20,1)", "SEARCHB(""~~"",J20,2)", "SEARCHB(""~~"",J20,3)", "SEARCHB(""~~"",J20,4)", "SEARCHB(""~~"",J20,5)", _
        "FIND("""",J20,1)", "FIND("""",J20,2)", "FIND("""",J20,3)", "FIND("""",J20,4)", "FIND("""",J20,5)", "FIND(K20,J20,1)", "FIND(K20,J20,2)", "FIND(K20,J20,3)", "FIND(K20,J20,4)", "FIND(K20,J20,5)", "FIND(L20,J20,1)", "FIND(L20,J20,2)", "FIND(L20,J20,3)", "FIND(L20,J20,4)", "FIND(L20,J20,5)", "FIND(""?"",J20,1)", _
        "FIND(""?"",J20,2)", "FIND(""?"",J20,3)", "FIND(""?"",J20,4)", "FIND(""?"",J20,5)", "FIND(""*"",J20,1)", "FIND(""*"",J20,2)", "FIND(""*"",J20,3)", "FIND(""*"",J20,4)", "FIND(""*"",J20,5)", "FIND(""~"",J20,1)", "FIND(""~"",J20,2)", "FIND(""~"",J20,3)", "FIND(""~"",J20,4)", "FIND(""~"",J20,5)", "FIND(""~x"",J20,1)", "FIND(""~x"",J20,2)", _
        "FIND(""~x"",J20,3)", "FIND(""~x"",J20,4)", "FIND(""~x"",J20,5)", "FIND(""~~"",J20,1)", "FIND(""~~"",J20,2)", "FIND(""~~"",J20,3)", "FIND(""~~"",J20,4)", "FIND(""~~"",J20,5)", "SEARCH("""",J20,1)", "SEARCH("""",J20,2)", "SEARCH("""",J20,3)", "SEARCH("""",J20,4)", "SEARCH("""",J20,5)", "SEARCH(K20,J20,1)", "SEARCH(K20,J20,2)", "SEARCH(K20,J20,3)", _
        "SEARCH(K20,J20,4)", "SEARCH(K20,J20,5)", "SEARCH(L20,J20,1)", "SEARCH(L20,J20,2)", "SEARCH(L20,J20,3)", "SEARCH(L20,J20,4)", "SEARCH(L20,J20,5)", "SEARCH(""?"",J20,1)", "SEARCH(""?"",J20,2)", "SEARCH(""?"",J20,3)", "SEARCH(""?"",J20,4)", "SEARCH(""?"",J20,5)", "SEARCH(""*"",J20,1)", "SEARCH(""*"",J20,2)", "SEARCH(""*"",J20,3)", "SEARCH(""*"",J20,4)", _
        "SEARCH(""*"",J20,5)", "SEARCH(""~"",J20,1)", "SEARCH(""~"",J20,2)", "SEARCH(""~"",J20,3)", "SEARCH(""~"",J20,4)", "SEARCH(""~"",J20,5)", "SEARCH(""~x"",J20,1)", "SEARCH(""~x"",J20,2)", "SEARCH(""~x"",J20,3)", "SEARCH(""~x"",J20,4)", "SEARCH(""~x"",J20,5)", "SEARCH(""~~"",J20,1)", "SEARCH(""~~"",J20,2)", "SEARCH(""~~"",J20,3)", "SEARCH(""~~"",J20,4)", "SEARCH(""~~"",J20,5)")
    Range("K20").Value2 = ChrW(&HD83D)
    Range("L20").Value2 = ChrW(&HDE00)
    For Each sample In samples
        If Len(report) > 0 Then report = report & "#"
        Range("J20").Value2 = sample
        report = report & UnitText(sample) & "="
        For Each expression In expressions
            value = Application.Evaluate("=" & expression)
            If IsError(value) Then
                report = report & "E:" & CStr(Application.Evaluate("=ERROR.TYPE(" & expression & ")")) & "|"
            Else
                report = report & ReportValue(value) & "|"
            End If
        Next expression
    Next sample
    ByteBoundaryControls = report
End Function
Public Function Driver() As Long

    Driver = 21

End Function


