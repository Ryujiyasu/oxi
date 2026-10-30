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
Public Trace As String
Public Number As Long
Public Sub BumpLong(ByRef value As Long)
    Trace = TypeName(value)
    value = value + 2
End Sub
Public Sub SetLarge(ByRef value As Long)
    Trace = "body"
    value = 40000
End Sub
Public Sub ChangeVariant(ByRef value As Variant)
    Trace = "body"
    value = "changed"
End Sub
Public Sub MutateRaise(ByRef value As Long)
    value = 12
    Trace = "body"
    Err.Raise 5
End Sub
Public Sub AliasLong(ByRef first As Long, ByRef second As Long)
    first = 7
    Trace = CStr(second)
    second = second + 1
End Sub
Public Sub AliasVariant(ByRef first As Variant, ByRef second As Variant)
    first = 7
    Trace = CStr(second)
    second = second + 1
End Sub
Public Sub ClearObject(ByRef value As Object)
    Set value = Nothing
    Trace = "body"
End Sub


Attribute VB_Name = "Driver"
Option Explicit
Private Function Case0() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case0 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case1() As String
    Dim obj As New Item
    Dim n As Integer
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case1 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case2() As String
    Dim obj As New Item
    Dim n As Double
    n = 40.5
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case2 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case3() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case3 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case4() As String
    Dim obj As New Item
    Dim n As Variant
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case4 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case5() As String
    Dim obj As New Item
    Dim n As Variant
    n = CLng(40)
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case5 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case6() As String
    Dim obj As New Item
    Dim n As Variant
    n = "40"
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case6 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case7() As String
    Dim obj As New Item
    Dim n As Integer
    n = 40
    On Error Resume Next
    CallByName obj, "SetLarge", VbMethod, n
    Case7 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case8() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "ChangeVariant", VbMethod, n
    Case8 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case9() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "MutateRaise", VbMethod, n
    Case9 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case10() As String
    Dim obj As New Item
    Dim n As Integer
    n = 40
    On Error Resume Next
    CallByName obj, "MutateRaise", VbMethod, n
    Case10 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case11() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "AliasLong", VbMethod, n, n
    Case11 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case12() As String
    Dim obj As New Item
    Dim n As Double
    n = 40.5
    On Error Resume Next
    CallByName obj, "AliasLong", VbMethod, n, n
    Case12 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case13() As String
    Dim obj As New Item
    Dim n As Variant
    n = 40
    On Error Resume Next
    CallByName obj, "AliasVariant", VbMethod, n, n
    Case13 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case14() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, (n)
    Case14 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case15() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n + 0
    Case15 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case16() As String
    Dim obj As New Item
    Dim n As Variant
    n = Null
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case16 = CStr(Err.Number) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case17() As String
    Dim obj As New Item
    Dim n As Variant
    n = CVErr(5)
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case17 = CStr(Err.Number) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case18() As String
    Dim obj As New Item
    Dim n(0 To 1) As Long
    n(0) = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n(0)
    Case18 = CStr(Err.Number) & "~" & CStr(n(0)) & "~" & TypeName(n(0)) & "~" & obj.Trace
End Function
Private Function Case19() As String
    Dim obj As New Item
    Dim holder As New Item
    holder.Number = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, holder.Number
    Case19 = CStr(Err.Number) & "~" & CStr(holder.Number) & "~" & TypeName(holder.Number) & "~" & obj.Trace
End Function
Private Function Case20() As String
    Dim obj As New Item
    Dim n As Object
    Set n = Range("A1")
    On Error Resume Next
    CallByName obj, "ClearObject", VbMethod, n
    Case20 = CStr(Err.Number) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case21() As String
    Dim obj As New Item
    Dim n As Range
    Set n = Range("A1")
    On Error Resume Next
    CallByName obj, "ClearObject", VbMethod, n
    Case21 = CStr(Err.Number) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Public Function Probe() As String
    Dim report As String
    report = report & "long=" & Case0() & vbLf
    report = report & "integer=" & Case1() & vbLf
    report = report & "double=" & Case2() & vbLf
    report = report & "string=" & Case3() & vbLf
    report = report & "variant_integer=" & Case4() & vbLf
    report = report & "variant_long=" & Case5() & vbLf
    report = report & "variant_string=" & Case6() & vbLf
    report = report & "integer_overflow=" & Case7() & vbLf
    report = report & "long_to_variant=" & Case8() & vbLf
    report = report & "long_error=" & Case9() & vbLf
    report = report & "integer_error=" & Case10() & vbLf
    report = report & "long_alias=" & Case11() & vbLf
    report = report & "double_alias=" & Case12() & vbLf
    report = report & "variant_alias=" & Case13() & vbLf
    report = report & "parentheses=" & Case14() & vbLf
    report = report & "expression=" & Case15() & vbLf
    report = report & "null=" & Case16() & vbLf
    report = report & "error_variant=" & Case17() & vbLf
    report = report & "array_element=" & Case18() & vbLf
    report = report & "class_field=" & Case19() & vbLf
    report = report & "object=" & Case20() & vbLf
    report = report & "range=" & Case21() & vbLf
    Probe = report
End Function
