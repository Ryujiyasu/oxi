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
    WidthProbeIdentity = 2026102011
End Function
Public Function WidthBoundary(ByVal requested As Double) As String
    Dim info As Variant, actual As Double, standard As Boolean
    Columns("D").UseStandardWidth = True
    Columns("D").ColumnWidth = requested
    actual = Columns("D").ColumnWidth
    standard = Columns("D").UseStandardWidth
    info = Application.Evaluate("CELL(""width"",D1)")
    If Not IsArray(info) Then
        WidthBoundary = "scalar:" & TypeName(info)
        Exit Function
    End If
    WidthBoundary = CStr(actual) & ":" & CStr(standard) & "|" & TypeName(info) & ":" & CStr(LBound(info)) & ":" & CStr(UBound(info))
    WidthBoundary = WidthBoundary & "|" & CStr(VarType(info(1))) & ":" & CStr(info(1)) & "|" & CStr(VarType(info(2))) & ":" & CStr(info(2))
End Function
