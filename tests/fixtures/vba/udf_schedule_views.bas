' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private scheduleKind As Long
Private scheduleCount As Long
Private scheduleLabels As String
Private scheduleStage As String

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102088
End Function

Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function

Private Function DescribeSched(ByVal value As Variant) As String
    DescribeSched = CStr(VarType(value)) & ":" & CStr(value)
End Function

Public Function SchedDouble(ByVal label As String, ByVal selected As Double) As Double
    Dim view As String
    If scheduleKind = 10 And label = "A" Then Application.Volatile
    If label = "B" Then view = DescribeSched(Range("M1").Value2)
    If label = "C" Then view = DescribeSched(Range("N1").Value2)
    scheduleCount = scheduleCount + 1
    scheduleLabels = scheduleLabels & scheduleStage & label & ":" & CStr(selected) & ":" & view & ":" & Application.Caller.Address(False, False) & ":" & CStr(Application.CalculationState) & ";"
    If label = "B" And (scheduleKind = 4 Or scheduleKind = 5 Or scheduleKind = 6 Or scheduleKind = 7) Then
        SchedDouble = 7
    Else
        SchedDouble = selected
    End If
End Function

Public Function SchedVariant(ByVal label As String, ByVal selected As Variant) As Double
    Dim value As Variant
    If IsObject(selected) Then
        value = selected.Value2
    Else
        value = selected
    End If
    scheduleCount = scheduleCount + 1
    scheduleLabels = scheduleLabels & scheduleStage & label & ":" & TypeName(selected) & ":" & DescribeSched(value) & ":" & Application.Caller.Address(False, False) & ":" & CStr(Application.CalculationState) & ";"
    If label = "B" And scheduleKind = 7 Then
        SchedVariant = 7
    Else
        SchedVariant = value
    End If
End Function

Private Function SnapSched(ByVal label As String) As String
    SnapSched = label & ":" & CStr(scheduleCount) & ":" & scheduleLabels
    SnapSched = SnapSched & "[" & DescribeSched(Range("M1").Value2) & "," & DescribeSched(Range("N1").Value2) & "," & DescribeSched(Range("O1").Value2) & "]"
End Function

Public Function ScheduleCase(ByVal kind As Long) As String
    Dim previous As Long, first As String, second As String, third As String, report As String
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:Q3").ClearContents
    Range("D1:D2").Value2 = 1
    If kind = 2 Then Range("D1").Value2 = 0
    If kind = 11 Then Range("D1").Value2 = 123
    scheduleKind = kind: scheduleCount = 0: scheduleLabels = "": scheduleStage = "E"
    If kind = 1 Then Range("Q1").Formula2 = "=SchedDouble(""Q"",D2)"
    first = "=SchedDouble(""A"",D1)"
    If kind = 12 Then
        Range("M1").Formula = first
    Else
        Range("M1").Formula2 = first
    End If
    If kind = 8 Then Range("Q1").Formula2 = "=SchedDouble(""Q"",D2)"
    If kind = 9 Or kind = 10 Then
        scheduleStage = "P"
        Application.Calculate
        scheduleStage = "E"
    End If
    second = "=SchedDouble(""B"",M1)"
    If kind = 3 Or kind = 7 Then second = "=SchedVariant(""B"",M1)"
    If kind = 13 Then
        ThisWorkbook.Names.Add Name:="ScheduleAlias", RefersTo:="='" & ActiveSheet.Name & "'!$M$1"
        second = "=SchedVariant(""B"",ScheduleAlias)"
    End If
    third = "=SchedDouble(""C"",N1)"
    If kind = 7 Then third = "=SchedVariant(""C"",N1)"
    If kind = 5 Then Range("O1").Formula2 = third
    If kind = 14 Then
        Range("N1:O1").Formula2 = second
    ElseIf kind = 12 Then
        Range("N1").Formula = second
    Else
        Range("N1").Formula2 = second
    End If
    If kind = 6 Or kind = 7 Then Range("O1").Formula2 = third
    report = SnapSched("entered")
    scheduleStage = "C"
    Application.Calculate
    report = report & "|" & SnapSched("calculated")
    If kind = 15 Then Range("D1").Value2 = 7
    scheduleStage = "D"
    Application.Calculate
    report = report & "|" & SnapSched("again")
    scheduleStage = "F"
    Application.CalculateFull
    report = report & "|" & SnapSched("full")
    scheduleStage = "R"
    Application.CalculateFullRebuild
    ScheduleCase = report & "|" & SnapSched("rebuilt")
    Range("M1:Q3").ClearContents
    If kind = 13 Then ThisWorkbook.Names("ScheduleAlias").Delete
    Application.Calculation = previous
End Function
