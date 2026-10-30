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

' SPDX-License-Identifier: CC-BY-SA-4.0
Attribute VB_Name = "SaveProbe"
Option Explicit
Public Function TypedCopyIdentity() As Long
    TypedCopyIdentity = 2026102033
End Function
Public Function CopyTyped(ByVal kind As Long, ByVal target As String, Optional ByVal dirty As Boolean = False) As String
    Dim candidate As Variant, number As Long, origin As String, description As String, help As String, context As Long
    Dim matrix(1 To 2, 1 To 2) As Variant, singleton(5 To 5) As String
    Range("A1").Value2 = target
    Range("A2").Value2 = target & ".cellA2.xlsm"
    Range("B1").Value2 = target & ".cellB1.xlsm"
    Select Case kind
    Case 0: candidate = Null
    Case 1: candidate = Empty
    Case 2: candidate = CVErr(2007)
    Case 3: candidate = Array(target)
    Case 4: Set candidate = Nothing
    Case 5: Set candidate = Range("A1")
    Case 6: candidate = target
    Case 7: candidate = vbNullString
    Case 8: Set candidate = New Collection
    Case 9: candidate = ""
    Case 10: candidate = Array(target, target & ".extra.xlsm")
    Case 11
        matrix(1, 1) = target
        matrix(1, 2) = target & ".matrix12.xlsm"
        matrix(2, 1) = target & ".matrix21.xlsm"
        matrix(2, 2) = target & ".matrix22.xlsm"
        candidate = matrix
    Case 12
        singleton(5) = target
        candidate = singleton
    Case 13: candidate = Array(Null)
    Case 14: candidate = Array()
    Case 15: candidate = Array(Array(target))
    Case 16: Set candidate = Range("A1:A2")
    Case 17: Set candidate = Range("A2:A1")
    Case 18
        Range("A1").ClearContents
        Set candidate = Range("A1")
    Case 19: Set candidate = Range("A1:B1")

    End Select
    ThisWorkbook.Saved = Not dirty
    On Error Resume Next
    Err.Clear
    ThisWorkbook.SaveCopyAs Filename:=candidate
    number = Err.Number
    origin = Err.Source
    description = Err.Description
    help = Err.HelpFile
    context = Err.HelpContext
    On Error GoTo 0
    CopyTyped = CStr(number) & "|" & origin & "|" & description & "|" & help & "|" & CStr(context) & "|saved:" & CStr(ThisWorkbook.Saved) & "|kind:" & TypeName(candidate)
End Function

