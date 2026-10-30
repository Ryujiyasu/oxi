' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private evaluationCalls As Long

' SPDX-License-Identifier: CC-BY-SA-4.0
Public Function StateProbeIdentity() As Long
    StateProbeIdentity = 2026101901
End Function

Public Function StateProbeTick() As Double
    StateProbeTick = Application.CalculationState
End Function

Private Function StateProbeSnapshot(ByVal twoSheets As Boolean) As String
    Dim report As String
    report = CStr(Application.CalculationState) & ":" & CStr(ThisWorkbook.Worksheets(1).Range("M1").Value2)
    If twoSheets Then report = report & ":" & CStr(ThisWorkbook.Worksheets(2).Range("M1").Value2)
    StateProbeSnapshot = report
End Function

Private Sub StateProbePrepare()
    Application.Calculation = xlCalculationManual
    If ThisWorkbook.Worksheets.Count < 2 Then
        ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If
    ThisWorkbook.Worksheets(1).Cells.ClearContents
    ThisWorkbook.Worksheets(2).Cells.ClearContents
    Application.CalculateFull
End Sub

Private Sub StateProbeAction(ByVal action As String)
    Select Case action
    Case "R": ThisWorkbook.Worksheets(1).Range("M1").Calculate
    Case "T": ThisWorkbook.Worksheets(2).Range("M1").Calculate
    Case "S": ThisWorkbook.Worksheets(1).Calculate
    Case "P": ThisWorkbook.Worksheets(2).Calculate
    Case "A": Application.Calculate
    Case "F": Application.CalculateFull
    Case "D": ThisWorkbook.Worksheets(1).Range("M1").Dirty
    Case "E": ThisWorkbook.Worksheets(2).Range("M1").Dirty
    Case "I": ThisWorkbook.Worksheets(1).Range("M1").Formula2 = "=StateProbeTick()"
    Case "J": ThisWorkbook.Worksheets(2).Range("M1").Formula2 = "=StateProbeTick()"
    Case "W": ThisWorkbook.Worksheets(1).Range("M1").ClearContents
    Case "K": ThisWorkbook.Worksheets(2).Range("M1").ClearContents
    Case "V": ThisWorkbook.Worksheets(1).Range("N1").Value2 = Application.Evaluate("StateProbeTick()")
    End Select
End Sub

Public Function StateProbeSteps(ByVal actions As String, ByVal automatic As Boolean, ByVal twoSheets As Boolean) As String
    Dim previousCalculation As Long, report As String, i As Long
    previousCalculation = Application.Calculation
    StateProbePrepare
    If automatic Then Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Worksheets(1).Range("M1").Formula2 = "=StateProbeTick()"
    If twoSheets Then ThisWorkbook.Worksheets(2).Range("M1").Formula2 = "=StateProbeTick()"
    report = StateProbeSnapshot(twoSheets)
    For i = 1 To Len(actions)
        StateProbeAction Mid$(actions, i, 1)
        report = report & "|" & StateProbeSnapshot(twoSheets)
    Next i
    StateProbeSteps = report
    StateProbePrepare
    Application.Calculation = previousCalculation
End Function

Public Function StateProbeBegin(ByVal twoSheets As Boolean) As String
    StateProbePrepare
    ThisWorkbook.Worksheets(1).Range("M1").Formula2 = "=StateProbeTick()"
    If twoSheets Then ThisWorkbook.Worksheets(2).Range("M1").Formula2 = "=StateProbeTick()"
    StateProbeBegin = StateProbeSnapshot(twoSheets)
End Function

Public Function StateProbeContinue(ByVal actions As String, ByVal twoSheets As Boolean) As String
    Dim report As String, i As Long
    report = StateProbeSnapshot(twoSheets)
    For i = 1 To Len(actions)
        StateProbeAction Mid$(actions, i, 1)
        report = report & "|" & StateProbeSnapshot(twoSheets)
    Next i
    StateProbeContinue = report
End Function

Public Function StateProbeConstants() As String
    StateProbeConstants = CStr(xlDone) & ":" & CStr(xlCalculating) & ":" & CStr(xlPending) & ":" & TypeName(Application.CalculationState)
End Function

' SPDX-License-Identifier: CC-BY-SA-4.0
Public Function StateProbeEvaluate(ByVal pending As Boolean, ByVal automatic As Boolean, ByVal worksheetScope As Boolean) As String
    Dim previousCalculation As Long, before As String, answer As Variant
    previousCalculation = Application.Calculation
    StateProbePrepare
    If automatic Then Application.Calculation = xlCalculationAutomatic
    If pending Then ThisWorkbook.Worksheets(1).Range("M1").Formula2 = "=StateProbeTick()"
    before = StateProbeSnapshot(False)
    If worksheetScope Then
        answer = ThisWorkbook.Worksheets(1).Evaluate("StateProbeTick()")
    Else
        answer = Application.Evaluate("StateProbeTick()")
    End If
    StateProbeEvaluate = before & "|E:" & CStr(answer) & "|" & StateProbeSnapshot(False)
    StateProbePrepare
    Application.Calculation = previousCalculation
End Function

Public Function StateProbeProperty() As String
    On Error Resume Next
    CallByName Application, "CalculationState", VbLet, 0
    StateProbeProperty = CStr(Err.Number) & ":" & CStr(Application.CalculationState)
    Err.Clear
End Function

Public Function StatusTick() As Double
    evaluationCalls = evaluationCalls + 1
    StatusTick = Application.CalculationState
End Function

Public Function CalculationStatusSteps(ByVal automatic As Boolean) As String
    Dim previousCalculation As Long, report As String, operations As String, i As Long
    previousCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:N5").ClearContents
    If automatic Then Application.Calculation = xlCalculationAutomatic
    Range("M1").Formula2 = "=StatusTick()"
    report = CStr(Application.CalculationState) & ":" & CStr(Range("M1").Value2)
    operations = "RASDAF"
    For i = 1 To Len(operations)
        Select Case Mid$(operations, i, 1)
        Case "R": Range("M1").Calculate
        Case "A": Application.Calculate
        Case "S": ActiveSheet.Calculate
        Case "D": Range("M1").Dirty
        Case "F": Application.CalculateFull
        End Select
        report = report & "|" & CStr(Application.CalculationState) & ":" & CStr(Range("M1").Value2)
    Next i
    CalculationStatusSteps = report
    Range("M1:N5").ClearContents
    Application.Calculation = previousCalculation
End Function

