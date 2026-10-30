' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private Function UnitText(ByVal text As String) As String
    Dim i As Long
    UnitText = CStr(Len(text))
    For i = 1 To Len(text)
        UnitText = UnitText & ":" & CStr(AscW(Mid$(text, i, 1)))
    Next i
End Function
Public Function SetupCells() As String
    Dim values As Variant
    Dim matrix(1 To 12, 1 To 1) As Variant
    Dim i As Long
    values = Array(ChrW(&HD83D), ChrW(&HDE00), ChrW(&HD83D) & ChrW(&HDE00), ChrW(&HD83D) & ChrW(0) & ChrW(&HDE00), ChrW(0), "_xD83D_", "_x005F_", "_x0000_", "A" & vbCr & "B", "A" & vbLf & "B", "<&>", "plain")
    For i = 1 To 12
        Cells(i, 1).Value = values(i - 1)
        Cells(i, 2).Value2 = values(i - 1)
        matrix(i, 1) = values(i - 1)
    Next i
    Range("C1:C12").Value2 = matrix
    Range("A1:A12").Copy Destination:=Range("D1")
    Range("C1:C12").Copy
    Range("E1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    SetupCells = ReadCells()
End Function
Public Function ReadCells() As String
    Dim row As Long
    Dim col As Long
    Dim report As String
    For row = 1 To 12
        If row > 1 Then report = report & "|"
        For col = 1 To 5
            If col > 1 Then report = report & "/"
            report = report & UnitText(Cells(row, col).Value2)
        Next col
    Next row
    ReadCells = report
End Function
Public Function Driver() As Long
    Driver = 21
End Function
