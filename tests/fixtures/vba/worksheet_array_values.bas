' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit

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
