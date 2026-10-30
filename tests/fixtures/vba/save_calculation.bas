' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at https://mozilla.org/MPL/2.0/.

Option Explicit

Public Sub Reset(ByVal useUdf As Boolean)
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateBeforeSave = True
    Range("A1").Value = 1
    If useUdf Then
        Range("B1").Formula = "=Twice(A1)"
    Else
        Range("B1").Formula = "=A1*2"
    End If
End Sub

Public Sub Configure(ByVal mode As Long, ByVal onSave As Boolean)
    Application.Calculation = mode
    Application.CalculateBeforeSave = onSave
    Range("A1").Value = 7
End Sub

Public Sub ModeOnly(ByVal mode As Long)
    Application.Calculation = mode
End Sub

Public Sub OnSaveOnly(ByVal onSave As Boolean)
    Application.CalculateBeforeSave = onSave
End Sub

Public Function Twice(ByVal amount As Double) As Double
    Twice = amount * 2
End Function
