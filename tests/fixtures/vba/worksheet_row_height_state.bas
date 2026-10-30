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
Public Function HeightProbeIdentity() As Long
    HeightProbeIdentity = 2026102025
End Function
Public Function HeightGeometry(ByVal requested As Double, ByVal zoom As Long, ByVal savePath As String) As String
    Dim observed As String, pixels As Long
    ActiveWindow.Zoom = zoom
    Rows("3:4").UseStandardHeight = True
    Rows("4").RowHeight = 19.5
    Rows("4").RowHeight = requested
    pixels = ActiveWindow.PointsToScreenPixelsY(72) - ActiveWindow.PointsToScreenPixelsY(0)
    observed = CStr(Rows("4").RowHeight) & "|" & CStr(Rows("4").Height) & "|" & CStr(Rows("4").Hidden) & "|" & CStr(Rows("4").UseStandardHeight) & "|" & CStr(ActiveWindow.Zoom) & "|" & CStr(pixels)
    ThisWorkbook.SaveCopyAs savePath
    HeightGeometry = observed
End Function

Public Function HeightBoundary(ByVal requested As Double, ByVal savePath As String) As String
    Dim failure As Long, observed As String
    Rows("3:4").UseStandardHeight = True
    Rows("4").RowHeight = 19.5
    On Error Resume Next
    Err.Clear
    Rows("4").RowHeight = requested
    failure = Err.Number
    On Error GoTo 0
    observed = CStr(failure) & "|" & CStr(Rows("4").RowHeight) & "|" & CStr(Rows("4").Height) & "|" & CStr(Rows("4").Hidden) & "|" & CStr(Rows("4").UseStandardHeight)
    ThisWorkbook.SaveCopyAs savePath
    HeightBoundary = observed
End Function
Public Function HeightStateChange(ByVal operation As Long, ByVal hidden As Boolean) As String
    Dim before As Variant, failure As Long, saved As Boolean, value As Variant
    Rows("3:4").UseStandardHeight = True
    Rows("4").RowHeight = 19.5
    Range("D3").Value2 = 1
    Range("D4").Value2 = 2
    Range("D5").Formula = "=SUBTOTAL(109,D3:D4)"
    Rows("4").Hidden = hidden
    before = Range("D5").Value2
    ThisWorkbook.Saved = True
    Set value = Nothing
    On Error Resume Next
    Err.Clear
    Select Case operation
    Case 0: Rows("4").UseStandardHeight = True
    Case 1: Rows("4").UseStandardHeight = False
    Case 2: Rows("4").UseStandardHeight = Null
    Case 3: Rows("4").UseStandardHeight = Empty
    Case 4: Rows("4").UseStandardHeight = CVErr(13)
    Case 5: Rows("4").RowHeight = 19.5
    Case 6: Rows("4").RowHeight = 0
    Case 7: Rows("4").RowHeight = 15
    Case 8: Rows("4").RowHeight = Null
    Case 9: Rows("4").RowHeight = Empty
    Case 10: Rows("4").RowHeight = 0.5
    Case 11: Rows("4").RowHeight = True
    Case 12: Rows("4").UseStandardHeight = value
    Case 13: Rows("4").RowHeight = 0.125
    Case 14: Rows("4").RowHeight = 409.50000001
    End Select
    failure = Err.Number
    saved = ThisWorkbook.Saved
    On Error GoTo 0
    HeightStateChange = CStr(failure) & "|" & CStr(saved) & "|" & CStr(before) & ">" & CStr(Range("D5").Value2) & "|" & CStr(Rows("4").RowHeight) & "|" & CStr(Rows("4").UseStandardHeight) & "|" & CStr(Rows("4").Hidden) & "|" & CStr(Application.Calculation)
End Function
