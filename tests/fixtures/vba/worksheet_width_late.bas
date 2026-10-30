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
    WidthProbeIdentity = 2026102018
End Function
Private Function ShowWidthValue(ByVal value As Variant) As String
    If IsNull(value) Then
        ShowWidthValue = "Null"
    Else
        ShowWidthValue = CStr(value)
    End If
End Function
Private Function WidthState() As String
    Dim info As Variant
    info = Application.Evaluate("CELL(""width"",D1)")
    WidthState = CStr(Columns("D").ColumnWidth) & ":" & CStr(Columns("D").UseStandardWidth) & ":" & CStr(ActiveSheet.StandardWidth) & ":" & CStr(Columns("D").Hidden)
    WidthState = WidthState & "|" & CStr(info(1)) & ":" & CStr(info(2))
End Function
Private Sub ResetWidth()
    ActiveSheet.StandardWidth = 8.43
    Columns("C:D").UseStandardWidth = True
    Columns("D").ColumnWidth = 12.5
End Sub
Public Function WidthCoercion(ByVal propertyName As String, ByVal kind As Long, Optional ByVal late As Boolean = False) As String
    Dim value As Variant, failure As Long
    ResetWidth
    Select Case kind
    Case 0: value = Empty
    Case 1: value = Null
    Case 2: value = True
    Case 3: value = False
    Case 4: value = 1.5
    Case 5: value = "1.5"
    Case 6: value = ""
    Case 7: value = "x"
    Case 8: value = CVErr(13)
    Case 9: value = Array(1, 2)
    Case 10: value = CDate(1.5)
    Case 11: value = CInt(2)
    Case 12: value = CCur(2.5)
    Case 13: value = CByte(2)
    Case 14: Set value = Nothing
    Case 15: Range("D1").Value2 = 1.5: Set value = Range("D1")
    Case 16: value = CDec(2.5)
    Case 17: value = "1,5"
    Case 18: value = "1,234"
    Case 19: value = 256
    Case 20: value = -1
    Case 21: value = "NaN"
    Case 22: value = CDate(0.125)
    Case 23: value = "True"
    Case 24: value = "False"
    End Select
    On Error Resume Next
    Err.Clear
    If late Then
        Select Case propertyName
        Case "column": CallByName Columns("D"), "ColumnWidth", VbLet, value
        Case "standard": CallByName ActiveSheet, "StandardWidth", VbLet, value
        Case "use": CallByName Columns("D"), "UseStandardWidth", VbLet, value
        End Select
    Else
    Select Case propertyName
    Case "column": Columns("D").ColumnWidth = value
    Case "standard": ActiveSheet.StandardWidth = value
    Case "use": Columns("D").UseStandardWidth = value
    End Select
    End If
    failure = Err.Number
    On Error GoTo 0
    WidthCoercion = CStr(failure) & "|" & WidthState()
End Function
Public Function WidthRestore(ByVal mode As Long) As String
    Dim before As String, hidden As String
    ResetWidth
    before = WidthState()
    Select Case mode
    Case 0: Columns("D").ColumnWidth = 0
    Case 1: Columns("D").ColumnWidth = 0.01
    Case 2: Columns("D").Hidden = True
    Case 3: ActiveSheet.StandardWidth = 0: Columns("D").UseStandardWidth = True
    Case 4: Columns("D").ColumnWidth = Empty
    Case 5: Columns("D").ColumnWidth = 0.49: Columns("D").Hidden = True
    End Select
    hidden = WidthState()
    Columns("D").Hidden = False
    WidthRestore = before & "|hidden:" & hidden & "|shown:" & WidthState()
End Function
Public Function WidthMixed(ByVal mode As Long) As String
    ActiveSheet.StandardWidth = 8.43
    Columns("C:D").UseStandardWidth = True
    Select Case mode
    Case 0: Columns("D").ColumnWidth = 8.43
    Case 1: Columns("D").ColumnWidth = 12.5
    Case 2: Columns("C:D").ColumnWidth = 8.43
    Case 3: Columns("D").Hidden = True
    Case 4: Columns("D").Hidden = False
    End Select
    WidthMixed = ShowWidthValue(Columns("C:D").ColumnWidth) & ":" & ShowWidthValue(Columns("C:D").UseStandardWidth) & "|" & WidthState()
End Function
