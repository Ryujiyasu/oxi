' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit

Public Sub ResetState()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateBeforeSave = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    Application.Cursor = xlDefault
    Application.Interactive = True
    Range("A1").Value = 1
    Range("B1").Formula = "=A1*2"
End Sub

Public Function ReadState() As String
    ReadState = CStr(Application.Calculation) & "|" & CStr(Application.ScreenUpdating) _
        & "|" & CStr(Application.EnableEvents) & "|" & CStr(Application.DisplayAlerts) _
        & "|" & CStr(VarType(Application.StatusBar)) & ":" & CStr(Application.StatusBar) _
        & "|" & CStr(Application.Cursor) & "|" & CStr(Application.Interactive) _
        & "|" & CStr(Range("B1").Value)
End Function

Public Function Configure(ByVal mode As Long) As String
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = "persist"
    Application.Cursor = xlWait
    Application.Interactive = False
    Range("A1").Value = 7
    Range("D1").Value = ReadState()
    Debug.Print ReadState()
    Select Case mode
        Case 1
            Error 11
        Case 2
            Error 40040
        Case 3
            On Error Resume Next
            Error 11
        Case 4
            Err.Raise 513, "own", "failure"
        Case 5
            Dim value As Variant
            value = Range("A1").Offset(-1, 0).Value
    End Select
    Configure = ReadState()
End Function

Public Function WriteMore() As String
    Range("A1").Value = 9
    WriteMore = ReadState()
End Function

Public Function Automatic() As String
    Application.Calculation = xlCalculationAutomatic
    Automatic = ReadState()
End Function

Public Function ResetText() As String
    Application.StatusBar = ""
    ResetText = ReadState()
End Function

Public Function SwitchMode(ByVal mode As Long) As String
    Application.Calculation = mode
    SwitchMode = ReadState()
End Function

Public Function ChangeInput(ByVal newValue As Long) As String
    Range("A1").Value = newValue
    ChangeInput = ReadState()
End Function
