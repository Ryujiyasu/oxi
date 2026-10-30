' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102135
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Public Function PhoneticCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, value As Variant, text As Variant, expression As String, code As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("D1:D5").ClearContents
    Select Case initial
        Case 0: text = "M1"
        Case 1: text = "ABC"
        Case 2: text = ChrW$(&H6771) & ChrW$(&H4EAC)
        Case 3: text = ChrW$(&H3068) & ChrW$(&H3046) & ChrW$(&H304D) & ChrW$(&H3087) & ChrW$(&H3046)
        Case 4: text = Empty
        Case 5: text = 1
    End Select
    Range("D2").Value2 = 0
    Range("D4").Value2 = text: Range("D5").Value2 = "other"
    Select Case kind
        Case 0: expression = "=PHONETIC(D4)"
        Case 1: expression = "=IF(TRUE,PHONETIC(D4),""other"")"
        Case 2: expression = "=IF(D2=0,PHONETIC(D4),PHONETIC(D5))"
        Case 3: expression = "=PHONETIC(IF(D2=0,D4,D5))"
        Case 4: expression = "=IFERROR(PHONETIC(D4),""fault"")"
        Case 5: expression = "=IF(D2=0,PHONETIC(D4)&"""",PHONETIC(D5)&"""")"
    End Select
    If kind < 6 Then
        Range("D1").Formula2 = expression
        Application.Calculate
        value = Range("D1").Value2
    ElseIf kind = 6 Then
        value = Application.Evaluate("PHONETIC(D4)")
    Else
        value = Application.WorksheetFunction.Phonetic(Range("D4"))
    End If
    If IsError(value) Then
        PhoneticCase = CStr(VarType(value)) & ":errorType=" & CStr(Application.WorksheetFunction.Error_Type(value))
    Else
        PhoneticCase = CStr(VarType(value)) & ":" & CStr(value)
    End If
    Range("D1:D5").ClearContents
    Application.Calculation = previous
    Exit Function
Failed:
    code = Err.Number: PhoneticCase = "fault@phonetic:" & CStr(code)
    On Error Resume Next
    Range("D1:D5").ClearContents
    Application.Calculation = previous
End Function
