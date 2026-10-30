' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit

Public Sub Seed()
    Application.StatusBar = False
    Worksheets("Sheet1").Range("A1").Value = 1
    Worksheets("Sheet1").Range("B1").Formula = "=A1*2"
End Sub

Private Sub BeforeFailure()
    Range("A1").Value = 7
    Range("C1").Value = "before"
    Range("A1").Font.Bold = True
    Range("A1").NumberFormat = "0.000"
    Application.StatusBar = "before-error"
    Debug.Print "before"
End Sub

Public Sub Success()
    BeforeFailure
    Range("C1").Value = "after"
End Sub

Public Sub SuccessNoWrite()
    Debug.Print "no-change"
End Sub

Public Sub DivideFailure()
    BeforeFailure
    Dim value As Double
    value = 1 / 0
    Range("C1").Value = "after"
End Sub

Public Sub RaiseFailure()
    BeforeFailure
    Err.Raise 513, "own", "failure"
    Range("C1").Value = "after"
End Sub

Public Sub AbortFailure()
    BeforeFailure
    Error 40040
    Range("C1").Value = "after"
End Sub

Public Sub ManualFailure()
    Application.Calculation = xlCalculationManual
    BeforeFailure
    Error 11
    Range("C1").Value = "after"
End Sub

Public Sub FormulaFailure()
    BeforeFailure
    Range("B1").Formula = "=A1*3"
    Error 11
    Range("C1").Value = "after"
End Sub

Public Sub ReadThenFailure()
    BeforeFailure
    Dim value As Variant
    value = Range("B1").Value
    Error 11
    Range("C1").Value = "after"
End Sub

Public Sub FormatFailure()
    Range("A1").Font.Bold = True
    Range("A1").NumberFormat = "0.000"
    Debug.Print "format"
    Error 11
End Sub

Public Sub StatusFailure()
    Application.StatusBar = "status-only"
    Debug.Print "status"
    Error 11
End Sub

Public Sub NoWriteFailure()
    Error 11
End Sub

Public Sub HostFailure()
    BeforeFailure
    Dim value As Variant
    value = Range("A1").Offset(-1, 0).Value
    Range("C1").Value = "after"
End Sub

Public Sub AddSheetFailure()
    BeforeFailure
    Dim added As Object
    Set added = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    added.Name = "Created"
    added.Range("D1").Value = "kept"
    Error 11
End Sub

Public Sub DeleteSheetFailure()
    BeforeFailure
    Application.DisplayAlerts = False
    Worksheets("Sheet2").Delete
    Error 11
End Sub

Public Sub RenameSheetFailure()
    BeforeFailure
    Worksheets("Sheet2").Name = "Renamed"
    Error 11
End Sub

Public Sub ActivateFailure()
    Worksheets("Sheet2").Activate
    Debug.Print "activated"
    Error 11
End Sub

Public Sub NameFailure()
    BeforeFailure
    ActiveWorkbook.Names.Add Name:="kept", RefersTo:="=Sheet1!$A$1"
    Error 11
End Sub

Public Sub HiddenRowFailure()
    BeforeFailure
    Rows(2).Hidden = True
    Error 11
End Sub

Public Sub ResumeNextSuccess()
    BeforeFailure
    On Error Resume Next
    Error 11
    Range("C1").Value = "after"
End Sub

Public Sub HandlerSuccess()
    On Error GoTo caught
    BeforeFailure
    Error 11
    Range("C1").Value = "after"
    Exit Sub
caught:
    Range("C1").Value = "caught"
End Sub

Public Sub NestedFailure()
    BeforeFailure
    InnerFailure
    Range("C1").Value = "after"
End Sub

Private Sub InnerFailure()
    Debug.Print "inner"
    Error 11
End Sub

Public Sub CompileEntryFailure()
    BeforeFailure
    Dim value As Variant
    value = CallByName(1, "Foo", VbGet)
End Sub

Public Sub CompileCalleeFailure()
    BeforeFailure
    BadCompile
    Range("C1").Value = "after"
End Sub

Private Sub BadCompile()
    Dim value As Variant
    value = CallByName(1, "Foo", VbGet)
End Sub
