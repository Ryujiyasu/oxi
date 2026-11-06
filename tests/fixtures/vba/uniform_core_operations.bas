Attribute VB_Name = "UniformCore"
' SPDX-License-Identifier: MPL-2.0
Option Explicit

Private Function Typed(ByVal value As Variant) As String
    Dim result As String, rank As Long, axis As Long
    Dim lower As Long, upper As Long, row As Long, col As Long
    result = CStr(VarType(value)) & ":" & TypeName(value) & ":"
    If IsArray(value) Then
        For axis = 1 To 4
            On Error Resume Next
            Err.Clear
            lower = LBound(value, axis)
            upper = UBound(value, axis)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
            rank = axis
            result = result & "[" & CStr(lower) & "," & CStr(upper) & "]"
        Next axis
        If rank = 1 Then
            For row = LBound(value) To UBound(value)
                result = result & "{" & Typed(value(row)) & "}"
            Next row
        ElseIf rank = 2 Then
            For row = LBound(value, 1) To UBound(value, 1)
                For col = LBound(value, 2) To UBound(value, 2)
                    result = result & "{" & Typed(value(row, col)) & "}"
                Next col
            Next row
        End If
    ElseIf IsError(value) Then
        result = result & CStr(value)
    ElseIf IsNull(value) Then
        result = result & "<null>"
    ElseIf IsEmpty(value) Then
        result = result & "<empty>"
    ElseIf VarType(value) = vbDate Then
        result = result & CStr(Year(value)) & "-" & CStr(Month(value)) & "-" & CStr(Day(value))
        result = result & "T" & CStr(Hour(value)) & ":" & CStr(Minute(value)) & ":" & CStr(Second(value))
    Else
        result = result & CStr(value)
    End If
    Typed = result
End Function

Public Function CaseProbe(ByVal kind As Long) As String
    Dim value As Variant, grid(1 To 2, 1 To 2) As Variant
    Dim report As String, target As Range, overlap As Range
    Dim items As Collection, dictionary As Object, held() As Variant
    On Error GoTo Failed
    Select Case kind
    Case 0
        CaseProbe = "oxi-uniform-core-20261106-v1"
        Exit Function
    Case 1
        value = CLng(7)
        report = Typed(value)
    Case 2
        report = Typed(value)
    Case 3
        value = Null
        report = Typed(value)
    Case 4
        value = CVErr(2042)
        report = Typed(value)
    Case 5
        value = Array(CLng(1), "two", Empty, True)
        report = Typed(value)
    Case 6
        grid(1, 1) = CLng(1): grid(1, 2) = "two"
        grid(2, 1) = Empty: grid(2, 2) = True
        value = grid
        report = Typed(value)
    Case 7
        Range("A1").Value2 = 3.25
        report = Typed(Range("A1").Value2)
    Case 8
        Range("A1").NumberFormat = "yyyy-mm-dd"
        Range("A1").Value2 = CLng(45292)
        report = Typed(Range("A1").Value) & "|" & Typed(Range("A1").Value2)
    Case 9
        Range("A1").NumberFormat = "$0.0000"
        Range("A1").Value2 = 12.3456789
        report = Typed(Range("A1").Value) & "|" & Typed(Range("A1").Value2)
    Case 10
        grid(1, 1) = CLng(1): grid(1, 2) = "two"
        grid(2, 1) = CVErr(2042): grid(2, 2) = True
        Range("A1:B2").Value2 = grid
        report = Typed(Range("A1:B2").Value2)
    Case 11
        Set target = Range("A1:B2").Offset(2, 3).Resize(3, 2)
        report = Typed(target.Address)
    Case 12
        report = Typed(Range("A1:C2").Rows.Count) & "|" & Typed(Range("A1:C2").Columns.Count)
        report = report & "|" & Typed(Range("A1:C2").Cells(2, 2).Address)
    Case 13
        Range("A1").Formula = "=1+2"
        Range("B1").Formula2 = "=SUM(3,4)"
        report = Typed(Range("A1").Formula) & "|" & Typed(Range("A1").Value2)
        report = report & "|" & Typed(Range("B1").Formula2) & "|" & Typed(Range("B1").Value2)
    Case 14
        Range("A1").Formula2 = "=SEQUENCE(2,2)"
        report = Typed(Range("A1:B2").Value2) & "|" & Typed(Range("A1").Formula2)
    Case 15
        report = Typed(Application.WorksheetFunction.Sum(Array(1, 2, 3)))
    Case 16
        report = Typed(Application.WorksheetFunction.Match("missing", Array("a", "b"), 0))
    Case 17
        report = Typed(Application.Evaluate("SUM(1,2)"))
    Case 18
        report = Typed(Application.Evaluate("{1,2;3,4}"))
    Case 19
        Set items = New Collection
        items.Add CLng(5), "first"
        items.Add "two", "second"
        report = Typed(items.Count) & "|" & Typed(items("first"))
        items.Remove "first"
        report = report & "|" & Typed(items.Count) & "|" & Typed(items(1))
    Case 20
        Set dictionary = CreateObject("Scripting.Dictionary")
        dictionary.CompareMode = vbTextCompare
        dictionary.Add "a", CLng(2)
        report = Typed(dictionary.Count) & "|" & Typed(dictionary.Exists("A")) & "|" & Typed(dictionary("A"))
    Case 21
        Range("A1").Value2 = "text"
        Range("A1").NumberFormat = "0.00"
        Range("A1").HorizontalAlignment = xlRight
        Range("A1").WrapText = True
        report = Typed(Range("A1").NumberFormat) & "|" & Typed(Range("A1").HorizontalAlignment)
        report = report & "|" & Typed(Range("A1").WrapText)
    Case 22
        Range("A1").Value2 = "Abg"
        Range("A1").Font.Size = 18
        Range("A1").Font.Bold = True
        report = Typed(Range("A1").Font.Size) & "|" & Typed(Range("A1").Font.Bold)
        report = report & "|" & Typed(Range("A1").RowHeight)
    Case 23
        Range("A1").Value2 = CLng(42)
        ThisWorkbook.Names.Add Name:="UniformName", RefersTo:="=Sheet1!$A$1"
        report = Typed(ThisWorkbook.Names("UniformName").RefersTo)
        report = report & "|" & Typed(ThisWorkbook.Names("UniformName").RefersToRange.Address)
    Case 24
        ThisWorkbook.Worksheets(1).Name = "Uniform Sheet"
        report = Typed(ThisWorkbook.Worksheets.Count) & "|" & Typed(ThisWorkbook.Worksheets(1).Name)
    Case 25
        Range("A1:B2").Value2 = CLng(7)
        Range("A1:B2").ClearContents
        report = Typed(Range("A1:B2").Value2) & "|" & Typed(Range("A1:B2").HasFormula)
    Case 26
        Set target = Application.Union(Range("A1:A2"), Range("C1:C2"))
        Set overlap = Application.Intersect(Range("A1:B2"), Range("B2:C3"))
        report = Typed(target.Address) & "|" & Typed(target.Areas.Count) & "|" & Typed(overlap.Address)
    Case 27
        Range("C3").Value2 = "x"
        report = Typed(ActiveSheet.UsedRange.Address)
    Case 28
        Range("A2").Value2 = CLng(7)
        Rows(2).Hidden = True
        report = Typed(Rows(2).Hidden) & "|" & Typed(Rows(2).RowHeight)
    Case 29
        Columns("B").ColumnWidth = 12.5
        Columns("B").Hidden = True
        report = Typed(Columns("B").Hidden) & "|" & Typed(Columns("B").ColumnWidth)
    Case 30
        ReDim held(-1 To 1)
        held(-1) = CLng(7): held(0) = "mid": held(1) = True
        ReDim Preserve held(-1 To 3)
        report = Typed(held)
    Case 31
        value = DateSerial(2024, 2, 29) + TimeSerial(13, 14, 15)
        report = Typed(value)
    Case 32
        value = ("Abc" Like "[A-C]*")
        report = Typed(value) & "|" & Typed(Split("a,b,c", ","))
    Case 33
        Range("A1:A2").NumberFormat = "$0.0000"
        Range("A1").Value = CCur(1.234567)
        Range("A2").Value2 = CCur(1.234567)
        report = Typed(Range("A1").Value2) & "|" & Typed(Range("A2").Value2)
    Case 34
        Range("A1").Value = DateSerial(2024, 2, 29)
        report = Typed(Range("A1").Value) & "|" & Typed(Range("A1").NumberFormat)
    Case 35
        Cells(2, 3).Value2 = CLng(9)
        report = Typed(Cells(2, 3).Value2) & "|" & Typed(Cells(2, 3).Address)
    Case Else
        Err.Raise 5, "UniformCore", "Unknown case"
    End Select
    CaseProbe = "ok|" & report
    Exit Function
Failed:
    CaseProbe = "error|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function StateReport() As String
    Dim sheet As Worksheet, named As String
    Set sheet = ThisWorkbook.Worksheets(1)
    On Error Resume Next
    named = ThisWorkbook.Names("UniformName").RefersTo
    If Err.Number <> 0 Then named = "<none>"
    Err.Clear
    On Error GoTo 0
    StateReport = sheet.Name & "|" & CStr(ThisWorkbook.Worksheets.Count)
    StateReport = StateReport & "|names=" & CStr(ThisWorkbook.Names.Count) & ":" & named
    StateReport = StateReport & "|format=" & CStr(sheet.Range("A1").NumberFormat)
    StateReport = StateReport & "|align=" & CStr(sheet.Range("A1").HorizontalAlignment)
    StateReport = StateReport & "|wrap=" & CStr(sheet.Range("A1").WrapText)
    StateReport = StateReport & "|font=" & CStr(sheet.Range("A1").Font.Name)
    StateReport = StateReport & ":" & CStr(sheet.Range("A1").Font.Size) & ":" & CStr(sheet.Range("A1").Font.Bold)
    StateReport = StateReport & "|row1=" & CStr(sheet.Rows(1).Hidden) & ":" & CStr(sheet.Rows(1).RowHeight)
    StateReport = StateReport & "|row2=" & CStr(sheet.Rows(2).Hidden) & ":" & CStr(sheet.Rows(2).RowHeight)
    StateReport = StateReport & "|colB=" & CStr(sheet.Columns("B").Hidden) & ":" & CStr(sheet.Columns("B").ColumnWidth)
End Function

Public Function NumberFormatProbe(ByVal pattern As String) As String
    On Error GoTo Failed
    Range("A1").Value2 = 12.3456789
    Range("A1").NumberFormat = pattern
    NumberFormatProbe = "ok|" & Range("A1").NumberFormat & "|" & Range("A1").NumberFormatLocal
    Exit Function
Failed:
    NumberFormatProbe = "error|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function NumberFormatRouteProbe(ByVal pattern As String, ByVal route As Long) As String
    Dim target As Object
    On Error GoTo Failed
    Set target = Range("A1")
    Range("A1").Value2 = 12.3456789
    Select Case route
    Case 0: Range("A1").NumberFormat = pattern
    Case 1: target.NumberFormat = pattern
    Case 2: CallByName target, "NumberFormat", VbLet, pattern
    Case 3: Range("A1").NumberFormatLocal = pattern
    Case 4: target.NumberFormatLocal = pattern
    Case 5: CallByName target, "NumberFormatLocal", VbLet, pattern
    End Select
    NumberFormatRouteProbe = "ok|" & Range("A1").NumberFormat & "|" & Range("A1").NumberFormatLocal
    Exit Function
Failed:
    NumberFormatRouteProbe = "error|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function
