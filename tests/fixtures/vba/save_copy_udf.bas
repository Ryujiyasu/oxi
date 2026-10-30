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
Private copyCalls As Long
Public Function CopyProbeIdentity() As Long
    CopyProbeIdentity = 2026102031
End Function
Public Function CopyTwice(ByVal value As Double) As Double
    copyCalls = copyCalls + 1
    CopyTwice = value * 2
End Function
Public Function CopyUdf(ByVal target As String, ByVal calculate As Boolean, ByVal saved As Boolean) As String
    Dim before As String, after As String, failure As Long
    Application.Calculation = -4105
    Range("B1").Value2 = 1
    Range("B2").Formula = "=CopyTwice(B1)"
    Application.Calculate
    Application.Calculation = -4135
    Range("B1").Value2 = 9
    Application.CalculateBeforeSave = calculate
    ThisWorkbook.Saved = saved
    copyCalls = 0
    before = CStr(Range("B2").Value2) & ":" & CStr(copyCalls) & ":" & CStr(ThisWorkbook.Saved) & ":" & CStr(Application.Calculation)
    On Error Resume Next
    Err.Clear
    ThisWorkbook.SaveCopyAs target
    failure = Err.Number
    On Error GoTo 0
    after = CStr(Range("B2").Value2) & ":" & CStr(copyCalls) & ":" & CStr(ThisWorkbook.Saved) & ":" & CStr(Application.Calculation)
    CopyUdf = CStr(failure) & "|" & before & "|" & after
End Function
