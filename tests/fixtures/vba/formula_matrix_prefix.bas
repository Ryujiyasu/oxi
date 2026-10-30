' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102147
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Private Function TypedValue(ByVal value As Variant) As String
    If IsEmpty(value) Then
        TypedValue = "0:Empty"
    ElseIf IsNull(value) Then
        TypedValue = "1:Null"
    Else
        TypedValue = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function
Private Function Snapshot(ByVal rowCount As Long, ByVal columnCount As Long) As String
    Dim r As Long, c As Long, cell As Range, answer As String
    For r = 1 To rowCount
        For c = 1 To columnCount
            Set cell = Cells(r, 13 + c)
            answer = answer & cell.Address(False, False) & "{v=" & TypedValue(cell.Value2)
            answer = answer & ";f=" & TypedValue(cell.Formula2) & ";r=" & TypedValue(cell.Formula2R1C1)
            answer = answer & ";has=" & CStr(cell.HasFormula) & "}"
        Next c
    Next r
    Snapshot = answer
End Function
Public Function MatrixCase(ByVal rowCount As Long, ByVal columnCount As Long, ByVal badIndex As Long, ByVal style As Long, ByVal negative As Long) As String
    Dim previous As Long, r As Long, c As Long, lowerRow As Long, lowerColumn As Long
    Dim matrix As Variant, target As Range, answer As String, fault As Long, calculateFault As Long
    On Error GoTo Failed
    previous = Application.Calculation: Application.Calculation = xlCalculationManual
    lowerRow = 1: lowerColumn = 1
    If negative <> 0 Then lowerRow = -2: lowerColumn = 5
    ReDim matrix(lowerRow To lowerRow + rowCount - 1, lowerColumn To lowerColumn + columnCount - 1)
    For r = 0 To rowCount - 1
        For c = 0 To columnCount - 1
            Cells(r + 1, c + 14).Value2 = 101 + 10 * r + c
            matrix(lowerRow + r, lowerColumn + c) = "=" & CStr(1 + 10 * r + c)
            If r * columnCount + c = badIndex Then matrix(lowerRow + r, lowerColumn + c) = "=PHONETIC(1)"
        Next c
    Next r
    Set target = Range(Cells(1, 14), Cells(rowCount, 13 + columnCount))
    answer = "before{" & Snapshot(rowCount, columnCount) & "}"
    On Error Resume Next
    Err.Clear
    If style = 0 Then
        target.Formula2 = matrix
    Else
        target.Formula2R1C1 = matrix
    End If
    fault = Err.Number: Err.Clear
    On Error GoTo Failed
    answer = answer & "|fault=" & CStr(fault) & "|immediate{" & Snapshot(rowCount, columnCount) & "}"
    On Error Resume Next
    Err.Clear: Application.Calculate: calculateFault = Err.Number: Err.Clear
    On Error GoTo Failed
    answer = answer & "|calculateFault=" & CStr(calculateFault) & "|calculated{" & Snapshot(rowCount, columnCount) & "}"
    Application.Calculation = previous: MatrixCase = answer
    Exit Function
Failed:
    fault = Err.Number: MatrixCase = "construction-fault=" & CStr(fault)
    On Error Resume Next
    Application.Calculation = previous
End Function
