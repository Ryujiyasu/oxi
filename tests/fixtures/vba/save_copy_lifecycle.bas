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
Public Function CopyProbeIdentity() As Long
    CopyProbeIdentity = 2026102029
End Function
Private Function CopyBookState() As String
    CopyBookState = ThisWorkbook.Name & "|" & ThisWorkbook.FullName & "|" & CStr(ThisWorkbook.Saved) & "|" & CStr(ThisWorkbook.ReadOnly) & "|" & CStr(ThisWorkbook.FileFormat)
End Function
Public Function CopyLifecycle(ByVal target As String, ByVal mode As Long) As String
    Dim before As String, after As String, failure As Long, origin As String, description As String, size As Long
    Range("B1").Value2 = mode
    ThisWorkbook.Saved = (mode <> 1)
    before = CopyBookState()
    On Error Resume Next
    Err.Clear
    ThisWorkbook.SaveCopyAs target
    failure = Err.Number
    origin = Err.Source
    description = Err.Description
    after = CopyBookState()
    If failure = 0 Then size = FileLen(target)
    On Error GoTo 0
    CopyLifecycle = CStr(failure) & "|" & origin & "|" & description & "|before:" & before & "|after:" & after & "|size:" & CStr(size)
End Function
