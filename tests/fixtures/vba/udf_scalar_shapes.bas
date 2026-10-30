' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hitsA As Long, hitsB As Long, hitsC As Long

Public Function FileAttributesIdentity() As Long
    FileAttributesIdentity = 2026102096
End Function
Public Function DirectoryEnvironment() As String
    DirectoryEnvironment = CurDir$
End Function
Public Function FlowTick(ByVal label As String, ByVal chosen As Double) As Double
    Select Case label
        Case "A": hitsA = hitsA + 1
        Case "B": hitsB = hitsB + 1
        Case "C": hitsC = hitsC + 1
    End Select
    FlowTick = chosen
End Function
Public Function ShapeVariant(ByVal item As Variant) As String
    Dim description As String
    hitsB = hitsB + 1
    description = CStr(VarType(item)) & "," & TypeName(item)
    description = description & "," & CStr(IsArray(item)) & "," & CStr(IsObject(item))
    If IsArray(item) Then
        On Error Resume Next
        description = description & "," & CStr(LBound(item, 1)) & "," & CStr(UBound(item, 1))
        description = description & "," & CStr(LBound(item, 2)) & "," & CStr(UBound(item, 2))
        description = description & ",err=" & CStr(Err.Number)
        On Error GoTo 0
    End If
    ShapeVariant = description
End Function
Public Function ShapeDouble(ByVal item As Double) As String
    hitsC = hitsC + 1
    ShapeDouble = CStr(VarType(item)) & "," & TypeName(item) & "," & CStr(item)
End Function
Private Function FlowSnapshot(ByVal stage As String) As String
    Dim m As Variant, n As Variant
    m = Range("M1").Value2: n = Range("N1").Value2
    FlowSnapshot = stage & ":" & CStr(hitsA) & "," & CStr(hitsB) & "," & CStr(hitsC)
    FlowSnapshot = FlowSnapshot & "[" & CStr(VarType(m)) & ":" & CStr(m) & "," & CStr(VarType(n)) & ":" & CStr(n) & "]"
End Function
Public Function FlowCase(ByVal kind As Long, ByVal initial As Long) As String
    Dim previous As Long, expression As String, report As String
    previous = Application.Calculation
    Application.Calculation = xlCalculationManual
    Range("M1:N2").ClearContents
    Range("D1").Value2 = initial
    hitsA = 0: hitsB = 0: hitsC = 0
    Range("M1").Formula2 = "=FlowTick(""A"",D1)"
    Select Case kind
        Case 0: expression = "=ShapeVariant(M1)"
        Case 1: expression = "=ShapeVariant((M1))"
        Case 2: expression = "=ShapeVariant(M1+0)"
        Case 3: expression = "=ShapeVariant(ROW(M1))"
        Case 4: expression = "=ShapeVariant(COLUMN(M1))"
        Case 5: expression = "=ShapeVariant({1})"
        Case 6: expression = "=ShapeVariant(SEQUENCE(1))"
        Case 7: expression = "=ShapeVariant(TRANSPOSE(M1))"
        Case 8: expression = "=ShapeDouble(M1)"
        Case 9: expression = "=ShapeDouble(ROW(M1))"
        Case 10: expression = "=ShapeDouble(COLUMN(M1))"
        Case 11: expression = "=ShapeDouble({1})"
        Case 12: expression = "=ShapeDouble(SEQUENCE(1))"
        Case 13: expression = "=ShapeDouble(TRANSPOSE(M1))"
        Case 14: expression = "=ShapeVariant(INDEX(M1,1))"
        Case 15: expression = "=ShapeVariant(OFFSET(M1,0,0))"
    End Select
    Range("N1").Formula2 = expression
    report = FlowSnapshot("entry")
    Application.Calculate
    report = report & "|" & FlowSnapshot("normal")
    Range("D1").Value2 = 1 - initial
    Application.Calculate
    report = report & "|" & FlowSnapshot("dirty")
    Application.CalculateFull
    FlowCase = report & "|" & FlowSnapshot("full")
    Range("M1:N2").ClearContents
    Application.Calculation = previous
End Function
