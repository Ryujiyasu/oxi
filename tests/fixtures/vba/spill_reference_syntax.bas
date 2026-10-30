' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102120
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Private Function RefAddress(ByVal text As String) As String
    Dim held As Range, code As Long
    On Error Resume Next
    Err.Clear
    Set held = Range(text)
    code = Err.Number
    If code <> 0 Then
        RefAddress = "error" & CStr(code)
    ElseIf held Is Nothing Then
        RefAddress = "Nothing"
    Else
        RefAddress = held.Address
    End If
End Function
Public Function ReferenceCase(ByVal kind As Long, ByVal query As Long) As String
    Dim previous As Long, requested As String, answer As String, code As Long
    On Error GoTo Failed
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:R4").ClearContents
    Select Case kind
        Case 0: Range("N1").Value2 = 1
        Case 1: Range("N1").ClearContents
        Case 2: Range("N1").Formula = "=1"
        Case 3: Range("N1").Formula2 = "=1"
        Case 4: Range("N1").Formula2 = "=SEQUENCE(2,2)"
    End Select
    ThisWorkbook.Names.Add Name:="SpillAnchor", RefersTo:="='" & ActiveSheet.Name & "'!$N$1"
    Application.CalculateFull
    Select Case query
        Case 0: requested = "N1#"
        Case 1: requested = "N2#"
        Case 2: requested = "SpillAnchor#"
        Case 3: requested = "N1#:Q2"
    End Select
    answer = RefAddress(requested)
    ThisWorkbook.Names("SpillAnchor").Delete
    Range("M1:R4").ClearContents
    Application.Calculation = previous
    ReferenceCase = answer
    Exit Function
Failed:
    code = Err.Number
    ReferenceCase = "error@reference:" & CStr(code)
    On Error Resume Next
    ThisWorkbook.Names("SpillAnchor").Delete
    Range("M1:R4").ClearContents
    Application.Calculation = previous
End Function
