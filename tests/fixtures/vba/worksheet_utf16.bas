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
Public Function Driver() As Long
    Driver = 21
End Function
