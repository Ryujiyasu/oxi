' SPDX-License-Identifier: CC-BY-SA-4.0
Public Function PersistByteFormulas() As String
    Range("J20").Value2 = "a" & ChrW(&HD83D) & ChrW(&HDE00) & "b"
    Range("K20").Value2 = ChrW(&HD83D)
    Range("L20").Value2 = ChrW(&HDE00)
    Range("M20").Formula = "=LEFTB(J20,3)"
    Range("N20").Formula = "=MIDB(J20,2,1)"
    Range("O20").Formula = "=REPLACEB(J20,3,1,K20)"
    Range("P20").Formula = "=FINDB(L20,J20,1)"
    Range("Q20").Formula = "=SEARCHB(""?"",J20,3)"
    Range("M20:Q20").Calculate
    PersistByteFormulas = ReadByteFormulas()
End Function

Public Function ReadByteFormulas() As String
    Dim address As Variant
    Dim report As String
    For Each address In Array("M20", "N20", "O20", "P20", "Q20")
        If Len(report) > 0 Then report = report & "|"
        report = report & ReportValue(Range(address).Value2)
    Next address
    ReadByteFormulas = report
End Function

Public Function ReadNamedByteFormulas(ByVal targetName As String) As String
    Dim address As Variant
    Dim report As String
    For Each address In Array("M20", "N20", "O20", "P20", "Q20")
        If Len(report) > 0 Then report = report & "|"
        report = report & ReportValue(Workbooks(targetName).Worksheets(1).Range(address).Value2)
    Next address
    ReadNamedByteFormulas = report
End Function
