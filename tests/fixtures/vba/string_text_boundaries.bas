' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private Type Holder
    text As String
End Type
Public Function TextCases() As String
    Dim high As String
    Dim low As String
    Dim items As Variant
    Dim report As String
    high = ChrW(&HD83D)
    low = ChrW(&HDE00)
    report = CStr(StrComp(high, low, vbTextCompare))
    report = report & "|" & CStr(InStr(1, high & low, low, vbTextCompare))
    report = report & "|" & UnitText(Replace(high & low, high, "x", 1, -1, vbTextCompare))
    items = Filter(Array(high, low), high, True, vbTextCompare)
    report = report & "|" & CStr(UBound(items)) & ":" & UnitText(Join(items, ""))
    report = report & "|" & CStr(StrPtr(UCase$(vbNullString)) = 0)
    report = report & "|" & CStr(StrPtr(Trim$(vbNullString)) = 0)
    TextCases = report
End Function
Public Function CellCases() As String
    Dim high As String
    Dim report As String
    high = ChrW(&HD83D)
    On Error Resume Next
    Range("A1").Value = high
    report = CStr(Err.Number)
    Err.Clear
    report = report & "|" & UnitText(Range("A1").Value)
    CellCases = report
End Function
Public Function Driver() As Long
    Driver = 21
End Function
Private Function UnitText(ByVal text As String) As String
    Dim i As Long
    Dim report As String
    report = CStr(Len(text)) & ":" & CStr(LenB(text))
    For i = 1 To Len(text)
        report = report & ":" & CStr(AscW(Mid$(text, i, 1)))
    Next i
    UnitText = report
End Function
Private Function ByValuePointer(ByVal text As String, ByVal pointer As LongLong) As Boolean
    ByValuePointer = (StrPtr(text) = pointer)
End Function
Private Function ByReferencePointer(ByRef text As String, ByVal pointer As LongLong) As Boolean
    ByReferencePointer = (StrPtr(text) = pointer)
End Function
Public Function PointerCases() As String
    Dim first As String
    Dim second As String
    Dim item As Holder
    Dim values(0 To 1) As String
    Dim pointer As LongLong
    Dim report As String
    report = CStr(StrPtr(first) = 0)
    first = ""
    report = report & "|" & CStr(StrPtr(first) = 0)
    first = vbNullString
    report = report & "|" & CStr(StrPtr(first) = 0)
    first = "cache"
    pointer = StrPtr(first)
    report = report & "|" & CStr(StrPtr(first) = pointer)
    second = first
    report = report & "|" & CStr(StrPtr(first) = StrPtr(second))
    report = report & "|" & CStr(ByValuePointer(first, pointer))
    report = report & "|" & CStr(ByReferencePointer(first, pointer))
    Mid$(first, 2, 1) = "x"
    report = report & "|" & CStr(StrPtr(first) = pointer)
    item.text = first
    report = report & "|" & CStr(StrPtr(item.text) = StrPtr(first))
    values(0) = first
    report = report & "|" & CStr(StrPtr(values(0)) = StrPtr(first))
    report = report & "|" & CStr(ByReferencePointer(values(0), StrPtr(values(0))))
    report = report & "|" & CStr(StrPtr(values(1)) = 0)
    report = report & "|" & CStr(StrPtr(vbNullString) = 0) & "|" & CStr(StrPtr("") = 0)
    PointerCases = report
End Function
Public Function UnicodeCases() As String
    Dim high As String
    Dim low As String
    Dim pair As String
    Dim parts As Variant
    Dim report As String
    high = ChrW(&HD83D)
    low = ChrW(&HDE00)
    pair = high & low
    report = UnitText(high) & "|" & UnitText(low) & "|" & UnitText(pair)
    report = report & "|" & UnitText(Left$(pair, 1)) & "|" & UnitText(Right$(pair, 1))
    report = report & "|" & UnitText(Mid$(pair, 1, 1))
    report = report & "|" & UnitText(StrReverse(pair))
    report = report & "|" & UnitText(Replace(pair, high, "x"))
    report = report & "|" & CStr(InStr(pair, low)) & ":" & CStr(StrComp(high, low, vbBinaryCompare))
    parts = Split(high & "|" & low, "|")
    report = report & "|" & UnitText(Join(parts, ""))
    report = report & "|" & UnitText(UCase$(high & "a")) & "|" & UnitText(Trim$(" " & high & " "))
    Mid$(pair, 1, 1) = "x"
    report = report & "|" & UnitText(pair)
    UnicodeCases = report
End Function
