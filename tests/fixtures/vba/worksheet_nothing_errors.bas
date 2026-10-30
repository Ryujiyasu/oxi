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
    WidthProbeIdentity = 2026102020
End Function
Public Function DimensionNothingError(ByVal propertyName As String, ByVal late As Boolean) As String
    Dim value As Variant, failure As Long, description As String, origin As String
    Dim helpFile As String, context As Long
    Set value = Nothing
    On Error Resume Next
    Err.Clear
    If late Then
        If propertyName = "StandardWidth" Then
            CallByName ActiveSheet, propertyName, VbLet, value
        Else
            CallByName Columns("D"), propertyName, VbLet, value
        End If
    Else
        Select Case propertyName
        Case "ColumnWidth": Columns("D").ColumnWidth = value
        Case "StandardWidth": ActiveSheet.StandardWidth = value
        Case "UseStandardWidth": Columns("D").UseStandardWidth = value
        Case "RowHeight": Columns("D").RowHeight = value
        Case "Hidden": Columns("D").Hidden = value
        Case "Value2": Columns("D").Value2 = value
        End Select
    End If
    failure = Err.Number
    description = Err.Description
    origin = Err.Source
    helpFile = Err.HelpFile
    context = Err.HelpContext
    On Error GoTo 0
    DimensionNothingError = CStr(failure) & "|" & origin & "|" & description & "|" & helpFile & "|" & CStr(context)
End Function
