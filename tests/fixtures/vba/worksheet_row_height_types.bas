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
    HeightProbeIdentity = 2026102022
End Function
Private Function ShowHeightValue(ByVal value As Variant) As String
    If IsNull(value) Then
        ShowHeightValue = "Null"
    Else
        ShowHeightValue = CStr(value)
    End If
End Function
Private Function HeightState() As String
    HeightState = ShowHeightValue(Rows("4").RowHeight) & ":" & ShowHeightValue(Rows("4").UseStandardHeight) & ":" & CStr(ActiveSheet.StandardHeight) & ":" & CStr(Rows("4").Hidden)
End Function
Public Function HeightCoercion(ByVal standard As Boolean, ByVal kind As Long) As String
    Dim value As Variant, failure As Long
    Rows("3:4").UseStandardHeight = True
    Rows("4").RowHeight = 19.5
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
    Case 19: value = 409.5
    Case 20: value = -1
    Case 21: value = "NaN"
    Case 22: value = CDate(0.125)
    Case 23: value = 409.50000001
    Case 24: value = 409.49999999
    Case 25: value = 0.374999999
    Case 26: value = 0.375
    Case 27: value = 0.749999999
    Case 28: value = 0.75
    Case 29: value = 19.49999999
    Case 30: value = 19.5
    Case 31: value = "True"
    Case 32: value = "False"
    End Select
    On Error Resume Next
    Err.Clear
    If standard Then
        Rows("4").UseStandardHeight = value
    Else
        Rows("4").RowHeight = value
    End If
    failure = Err.Number
    On Error GoTo 0
    HeightCoercion = CStr(failure) & "|" & HeightState()
End Function
