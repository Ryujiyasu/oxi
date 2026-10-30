VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Echo(ByVal value As Variant) As String
    Driver.Trace = Driver.Trace & "B"
    Echo = "ok"
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Private heldObject As Item
Private Function GetTarget(ByVal mode As Long) As Variant
    Trace = Trace & "O"
    If mode = 1 Then
        GetTarget = 42
    ElseIf mode = 2 Then
        Set GetTarget = Nothing
    Else
        Set GetTarget = heldObject
    End If
End Function
Private Function GetMemberName(ByVal value As Variant) As Variant
    Trace = Trace & "N"
    GetMemberName = value
End Function
Private Function GetInvocation(ByVal value As Variant) As Variant
    Trace = Trace & "K"
    GetInvocation = value
End Function
Private Function GetRhs() As Variant
    Trace = Trace & "R"
    GetRhs = "value"
End Function
Private Function Case0() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(CVErr(5)), GetInvocation(1), GetRhs())
    Case0 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case1() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(Null), GetInvocation(1), GetRhs())
    Case1 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case2() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(Array(1)), GetInvocation(1), GetRhs())
    Case2 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case3() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(Empty), GetInvocation(1), GetRhs())
    Case3 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case4() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(CVErr(5)), GetRhs())
    Case4 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case5() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(Array(1)), GetRhs())
    Case5 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case6() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(Null), GetRhs())
    Case6 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case7() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation("bad"), GetRhs())
    Case7 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case8() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(2147483648#), GetRhs())
    Case8 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case9() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName(CVErr(5)), GetInvocation(1), GetRhs())
    Case9 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case10() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName(Null), GetInvocation(1), GetRhs())
    Case10 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case11() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName(Array(1)), GetInvocation(1), GetRhs())
    Case11 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case12() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName(Empty), GetInvocation(1), GetRhs())
    Case12 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case13() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName("Echo"), GetInvocation(CVErr(5)), GetRhs())
    Case13 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case14() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName("Echo"), GetInvocation(Array(1)), GetRhs())
    Case14 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case15() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName("Echo"), GetInvocation(Null), GetRhs())
    Case15 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case16() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName("Echo"), GetInvocation("bad"), GetRhs())
    Case16 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case17() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName("Echo"), GetInvocation(2147483648#), GetRhs())
    Case17 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case18() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName(CVErr(5)), GetInvocation(1), GetRhs())
    Case18 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case19() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName(Null), GetInvocation(1), GetRhs())
    Case19 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case20() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName(Array(1)), GetInvocation(1), GetRhs())
    Case20 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case21() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName(Empty), GetInvocation(1), GetRhs())
    Case21 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case22() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName("Echo"), GetInvocation(CVErr(5)), GetRhs())
    Case22 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case23() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName("Echo"), GetInvocation(Array(1)), GetRhs())
    Case23 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case24() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName("Echo"), GetInvocation(Null), GetRhs())
    Case24 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case25() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName("Echo"), GetInvocation("bad"), GetRhs())
    Case25 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case26() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName("Echo"), GetInvocation(2147483648#), GetRhs())
    Case26 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case27() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(Null), GetInvocation("bad"), GetRhs())
    Case27 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case28() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(CVErr(5)), GetInvocation(Null), GetRhs())
    Case28 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case29() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Absent"), GetInvocation("bad"), GetRhs())
    Case29 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case30() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName(Null), GetInvocation(2147483648#), GetRhs())
    Case30 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Public Function Probe() As String
    Dim report As String
    report = report & "object_error=" & Case0() & vbLf
    report = report & "object_null=" & Case1() & vbLf
    report = report & "object_array=" & Case2() & vbLf
    report = report & "object_empty=" & Case3() & vbLf
    report = report & "object_kind_error=" & Case4() & vbLf
    report = report & "object_kind_array=" & Case5() & vbLf
    report = report & "object_kind_null=" & Case6() & vbLf
    report = report & "object_kind_bad=" & Case7() & vbLf
    report = report & "object_kind_overflow=" & Case8() & vbLf
    report = report & "scalar_error=" & Case9() & vbLf
    report = report & "scalar_null=" & Case10() & vbLf
    report = report & "scalar_array=" & Case11() & vbLf
    report = report & "scalar_empty=" & Case12() & vbLf
    report = report & "scalar_kind_error=" & Case13() & vbLf
    report = report & "scalar_kind_array=" & Case14() & vbLf
    report = report & "scalar_kind_null=" & Case15() & vbLf
    report = report & "scalar_kind_bad=" & Case16() & vbLf
    report = report & "scalar_kind_overflow=" & Case17() & vbLf
    report = report & "nothing_error=" & Case18() & vbLf
    report = report & "nothing_null=" & Case19() & vbLf
    report = report & "nothing_array=" & Case20() & vbLf
    report = report & "nothing_empty=" & Case21() & vbLf
    report = report & "nothing_kind_error=" & Case22() & vbLf
    report = report & "nothing_kind_array=" & Case23() & vbLf
    report = report & "nothing_kind_null=" & Case24() & vbLf
    report = report & "nothing_kind_bad=" & Case25() & vbLf
    report = report & "nothing_kind_overflow=" & Case26() & vbLf
    report = report & "null_bad=" & Case27() & vbLf
    report = report & "error_null=" & Case28() & vbLf
    report = report & "absent_bad=" & Case29() & vbLf
    report = report & "null_overflow=" & Case30() & vbLf
    Probe = report
End Function
