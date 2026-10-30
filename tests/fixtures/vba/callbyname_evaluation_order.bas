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
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(1), GetRhs())
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
    result = CallByName(GetTarget(0), GetMemberName(CVErr(5)), GetInvocation(1), GetRhs())
    Case2 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case3() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(Null), GetRhs())
    Case3 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case4() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation("bad"), GetRhs())
    Case4 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case5() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(2147483648#), GetRhs())
    Case5 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case6() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(1), GetMemberName("Echo"), GetInvocation(1), GetRhs())
    Case6 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case7() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(2), GetMemberName("Echo"), GetInvocation(1), GetRhs())
    Case7 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case8() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Absent"), GetInvocation(1), GetRhs())
    Case8 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Private Function Case9() As String
    Dim result As Variant
    Set heldObject = New Item
    Trace = ""
    On Error Resume Next
    result = CallByName(GetTarget(0), GetMemberName("Echo"), GetInvocation(16), GetRhs())
    Case9 = CStr(Err.Number) & "~" & Trace & "~" & CStr(result)
End Function
Public Function Probe() As String
    Dim report As String
    report = report & "valid=" & Case0() & vbLf
    report = report & "null_name=" & Case1() & vbLf
    report = report & "error_name=" & Case2() & vbLf
    report = report & "null_kind=" & Case3() & vbLf
    report = report & "bad_kind=" & Case4() & vbLf
    report = report & "overflow_kind=" & Case5() & vbLf
    report = report & "scalar_target=" & Case6() & vbLf
    report = report & "nothing_target=" & Case7() & vbLf
    report = report & "missing_member=" & Case8() & vbLf
    report = report & "invalid_flags=" & Case9() & vbLf
    Probe = report
End Function
