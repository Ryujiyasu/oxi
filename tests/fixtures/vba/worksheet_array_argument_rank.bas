VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "SaveProbe"
' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function WidthProbeIdentity() As Long
    WidthProbeIdentity = 2026102012
End Function
Public Function RankRoot(ByVal expression As String) As String
    Dim answer As Variant
    answer = Application.Evaluate(expression)
    RankRoot = ShapeAndValue(answer)
End Function
Public Function RankArgument(ByVal answer As Variant) As String
    RankArgument = ShapeAndValue(answer)
End Function
Public Function RankFormula(ByVal expression As String, ByVal legacy As Boolean) As String
    Dim previous As Long
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:Q5").ClearContents
    If legacy Then
        Range("M1").Formula = "=RankArgument(" & expression & ")"
    Else
        Range("M1").Formula2 = "=RankArgument(" & expression & ")"
    End If
    RankFormula = CStr(Range("M1").Value2)
    Range("M1").Calculate
    RankFormula = RankFormula & "|again:" & CStr(Range("M1").Value2)
    Range("M1:Q5").ClearContents
    Application.Calculation = previous
End Function
Private Function ScalarReport(ByVal value As Variant) As String
    Dim i As Long
    If IsError(value) Then
        ScalarReport = "E:" & CStr(value)
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
