' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private eventKind As Long
Private eventCount As Long
Private events As String
Private eventStage As String

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102089
End Function

Public Function DirectoryEnvironment() As String
    Dim report As String, failure As Long
    report = CurDir$
    On Error Resume Next
    Err.Clear
    report = report & "|mtc:" & CStr(Application.MultiThreadedCalculation.Enabled)
    report = report & "|threads:" & CStr(Application.MultiThreadedCalculation.ThreadCount)
    failure = Err.Number
    On Error GoTo 0
    DirectoryEnvironment = report & "|metadata_error:" & CStr(failure)
End Function

Public Function EarlyTick(ByVal label As String, ByVal selected As Double, ByVal readView As Boolean) As Double
    Dim callId As Long, view As Variant
    eventCount = eventCount + 1: callId = eventCount
    events = events & "+" & eventStage & label & ":" & CStr(callId) & ":" & CStr(selected) & ";"
    If readView Then
        If label = "B" Then view = Range("M1").Value2
        If label = "C" Then view = Range("N1").Value2
    End If
    events = events & "-" & eventStage & label & ":" & CStr(callId) & ":" & CStr(VarType(view)) & ":" & CStr(view) & ";"
    If label = "B" And eventKind >= 6 Then
        EarlyTick = 7
    Else
        EarlyTick = selected
    End If
End Function

Public Function EventCase(ByVal kind As Long) As String
    Dim previous As Long, readText As String, report As String
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:Q3").ClearContents
    Range("D1:D2").Value2 = 1
    eventKind = kind: eventCount = 0: events = "": eventStage = "E"
    If kind Mod 2 = 1 Then readText = "TRUE" Else readText = "FALSE"
    If kind = 2 Or kind = 3 Then Range("Q1").Formula2 = "=EarlyTick(""Q"",D2," & readText & ")"
    Range("M1").Formula2 = "=EarlyTick(""A"",D1," & readText & ")"
    If kind = 4 Or kind = 5 Then Range("Q1").Formula2 = "=EarlyTick(""Q"",D2," & readText & ")"
    Range("N1").Formula2 = "=EarlyTick(""B"",M1," & readText & ")"
    If kind >= 6 Then Range("O1").Formula2 = "=EarlyTick(""C"",N1," & readText & ")"
    report = "entered:" & CStr(eventCount) & ":" & events
    eventStage = "C"
    Application.Calculate
    report = report & "|calculated:" & CStr(eventCount) & ":" & events
    eventStage = "D"
    Range("D1").Value2 = 7
    Application.Calculate
    report = report & "|dirty:" & CStr(eventCount) & ":" & events
    eventStage = "F"
    Application.CalculateFull
    EventCase = report & "|full:" & CStr(eventCount) & ":" & events
    Range("M1:Q3").ClearContents
    Application.Calculation = previous
End Function
