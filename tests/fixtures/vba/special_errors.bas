' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.
Option Explicit
Private progress As String

Public Function StatementResumeNext(ByVal number As Long) As String
    On Error Resume Next
    progress = "before"
    Error number
    progress = "after"
    StatementResumeNext = Snapshot()
End Function

Public Function StatementGoTo(ByVal number As Long) As String
    On Error GoTo caught
    progress = "before"
    Error number
    progress = "after"
    StatementGoTo = Snapshot()
    Exit Function
caught:
    progress = "handler"
    StatementGoTo = Snapshot()
End Function

Public Function StatementUnhandled(ByVal number As Long) As String
    progress = "before"
    Error number
    progress = "after"
    StatementUnhandled = Snapshot()
End Function

Public Function StatementNested(ByVal number As Long) As String
    On Error Resume Next
    progress = "before"
    InnerError number
    progress = progress & "|outer-after"
    StatementNested = Snapshot()
End Function

Private Sub InnerError(ByVal number As Long)
    On Error Resume Next
    Error number
    progress = progress & "|inner-after"
End Sub

Public Function RaiseResumeNext(ByVal number As Long, Optional description As Variant) As String
    On Error Resume Next
    progress = "before"
    If IsMissing(description) Then
        Err.Raise number
    Else
        Err.Raise number, "custom", description
    End If
    progress = "after"
    RaiseResumeNext = Snapshot()
End Function

Public Function RaiseGoTo(ByVal number As Long, Optional description As Variant) As String
    On Error GoTo caught
    progress = "before"
    If IsMissing(description) Then
        Err.Raise number
    Else
        Err.Raise number, "custom", description
    End If
    progress = "after"
    RaiseGoTo = Snapshot()
    Exit Function
caught:
    progress = "handler"
    RaiseGoTo = Snapshot()
End Function

Public Function RaiseUnhandled(ByVal number As Long, Optional description As Variant) As String
    progress = "before"
    If IsMissing(description) Then
        Err.Raise number
    Else
        Err.Raise number, "custom", description
    End If
    progress = "after"
    RaiseUnhandled = Snapshot()
End Function

Public Function AssignNumber(ByVal number As Long) As String
    On Error Resume Next
    Err.Number = number
    progress = "after"
    AssignNumber = Snapshot()
End Function

Public Function RaiseWithState(ByVal number As Long) As String
    On Error Resume Next
    Err.Number = 6
    Err.Source = "old"
    Err.Description = "old description"
    Err.HelpFile = "oldHelp.chm"
    Err.HelpContext = 77
    Err.Raise number
    RaiseWithState = Snapshot()
End Function

Public Function DispatchRaise(ByVal number As Long) As String
    Dim target As Object, ignored As Variant
    Set target = Err
    On Error Resume Next
    ignored = CallByName(target, "Raise", VbMethod, number)
    DispatchRaise = Snapshot()
End Function

Public Function ReadProgress() As String
    ReadProgress = progress
End Function

Public Function Sweep(ByVal useRaise As Boolean) As Variant
    Dim values(0 To 65535) As String, code As Long
    On Error Resume Next
    For code = 0 To 65535
        If useRaise Or code <> 40040 Then
            Err.Clear
            If useRaise Then
                Err.Raise code
            Else
                Error code
            End If
            values(code) = Err.Number & "|" & Err.Description & "|" & Err.HelpContext
        Else
            values(code) = "ABORT"
        End If
    Next code
    Sweep = values
End Function

Private Function Snapshot() As String
    Snapshot = Err.Number & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & Err.HelpContext
End Function
