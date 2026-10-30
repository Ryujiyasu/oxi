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
Public Sub ChangeByte(ByRef value As Byte)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeInteger(ByRef value As Integer)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeLong(ByRef value As Long)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeLongLong(ByRef value As LongLong)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeLongPtr(ByRef value As LongPtr)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeSingle(ByRef value As Single)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeDouble(ByRef value As Double)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeCurrency(ByRef value As Currency)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeDate(ByRef value As Date)
    Trace = TypeName(value)
    value = value + 1
End Sub
Public Sub ChangeBoolean(ByRef value As Boolean)
    Trace = TypeName(value)
    value = Not value
End Sub
Public Sub ChangeString(ByRef value As String)
    Trace = TypeName(value)
    value = value & "x"
End Sub


Attribute VB_Name = "Driver"
Option Explicit
Private Function Case0() As String
    Dim obj As New Item
    Dim n As Byte
    n = CByte(40)
    On Error Resume Next
    CallByName obj, "ChangeByte", VbMethod, n
    Case0 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case1() As String
    Dim obj As New Item
    Dim n As Variant
    n = CByte(40)
    On Error Resume Next
    CallByName obj, "ChangeByte", VbMethod, n
    Case1 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case2() As String
    Dim obj As New Item
    Dim n As Byte
    n = CByte(40)
    On Error Resume Next
    CallByName obj, "ChangeByte", VbMethod, (n)
    Case2 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case3() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeByte", VbMethod, n
    Case3 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case4() As String
    Dim obj As New Item
    Dim n As Integer
    n = CInt(40)
    On Error Resume Next
    CallByName obj, "ChangeInteger", VbMethod, n
    Case4 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case5() As String
    Dim obj As New Item
    Dim n As Variant
    n = CInt(40)
    On Error Resume Next
    CallByName obj, "ChangeInteger", VbMethod, n
    Case5 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case6() As String
    Dim obj As New Item
    Dim n As Integer
    n = CInt(40)
    On Error Resume Next
    CallByName obj, "ChangeInteger", VbMethod, (n)
    Case6 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case7() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeInteger", VbMethod, n
    Case7 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case8() As String
    Dim obj As New Item
    Dim n As Long
    n = CLng(40)
    On Error Resume Next
    CallByName obj, "ChangeLong", VbMethod, n
    Case8 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case9() As String
    Dim obj As New Item
    Dim n As Variant
    n = CLng(40)
    On Error Resume Next
    CallByName obj, "ChangeLong", VbMethod, n
    Case9 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case10() As String
    Dim obj As New Item
    Dim n As Long
    n = CLng(40)
    On Error Resume Next
    CallByName obj, "ChangeLong", VbMethod, (n)
    Case10 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case11() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeLong", VbMethod, n
    Case11 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case12() As String
    Dim obj As New Item
    Dim n As LongLong
    n = CLngLng(40)
    On Error Resume Next
    CallByName obj, "ChangeLongLong", VbMethod, n
    Case12 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case13() As String
    Dim obj As New Item
    Dim n As Variant
    n = CLngLng(40)
    On Error Resume Next
    CallByName obj, "ChangeLongLong", VbMethod, n
    Case13 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case14() As String
    Dim obj As New Item
    Dim n As LongLong
    n = CLngLng(40)
    On Error Resume Next
    CallByName obj, "ChangeLongLong", VbMethod, (n)
    Case14 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case15() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeLongLong", VbMethod, n
    Case15 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case16() As String
    Dim obj As New Item
    Dim n As LongPtr
    n = CLngPtr(40)
    On Error Resume Next
    CallByName obj, "ChangeLongPtr", VbMethod, n
    Case16 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case17() As String
    Dim obj As New Item
    Dim n As Variant
    n = CLngPtr(40)
    On Error Resume Next
    CallByName obj, "ChangeLongPtr", VbMethod, n
    Case17 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case18() As String
    Dim obj As New Item
    Dim n As LongPtr
    n = CLngPtr(40)
    On Error Resume Next
    CallByName obj, "ChangeLongPtr", VbMethod, (n)
    Case18 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case19() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeLongPtr", VbMethod, n
    Case19 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case20() As String
    Dim obj As New Item
    Dim n As Single
    n = CSng(40.5)
    On Error Resume Next
    CallByName obj, "ChangeSingle", VbMethod, n
    Case20 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case21() As String
    Dim obj As New Item
    Dim n As Variant
    n = CSng(40.5)
    On Error Resume Next
    CallByName obj, "ChangeSingle", VbMethod, n
    Case21 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case22() As String
    Dim obj As New Item
    Dim n As Single
    n = CSng(40.5)
    On Error Resume Next
    CallByName obj, "ChangeSingle", VbMethod, (n)
    Case22 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case23() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeSingle", VbMethod, n
    Case23 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case24() As String
    Dim obj As New Item
    Dim n As Double
    n = CDbl(40.5)
    On Error Resume Next
    CallByName obj, "ChangeDouble", VbMethod, n
    Case24 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case25() As String
    Dim obj As New Item
    Dim n As Variant
    n = CDbl(40.5)
    On Error Resume Next
    CallByName obj, "ChangeDouble", VbMethod, n
    Case25 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case26() As String
    Dim obj As New Item
    Dim n As Double
    n = CDbl(40.5)
    On Error Resume Next
    CallByName obj, "ChangeDouble", VbMethod, (n)
    Case26 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case27() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeDouble", VbMethod, n
    Case27 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case28() As String
    Dim obj As New Item
    Dim n As Currency
    n = CCur(40.5)
    On Error Resume Next
    CallByName obj, "ChangeCurrency", VbMethod, n
    Case28 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case29() As String
    Dim obj As New Item
    Dim n As Variant
    n = CCur(40.5)
    On Error Resume Next
    CallByName obj, "ChangeCurrency", VbMethod, n
    Case29 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case30() As String
    Dim obj As New Item
    Dim n As Currency
    n = CCur(40.5)
    On Error Resume Next
    CallByName obj, "ChangeCurrency", VbMethod, (n)
    Case30 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case31() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeCurrency", VbMethod, n
    Case31 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case32() As String
    Dim obj As New Item
    Dim n As Date
    n = CDate(DateSerial(2020, 1, 2))
    On Error Resume Next
    CallByName obj, "ChangeDate", VbMethod, n
    Case32 = CStr(Err.Number) & "~" & CStr(CDbl(n)) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case33() As String
    Dim obj As New Item
    Dim n As Variant
    n = CDate(DateSerial(2020, 1, 2))
    On Error Resume Next
    CallByName obj, "ChangeDate", VbMethod, n
    Case33 = CStr(Err.Number) & "~" & CStr(CDbl(n)) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case34() As String
    Dim obj As New Item
    Dim n As Date
    n = CDate(DateSerial(2020, 1, 2))
    On Error Resume Next
    CallByName obj, "ChangeDate", VbMethod, (n)
    Case34 = CStr(Err.Number) & "~" & CStr(CDbl(n)) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case35() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeDate", VbMethod, n
    Case35 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case36() As String
    Dim obj As New Item
    Dim n As Boolean
    n = CBool(-1)
    On Error Resume Next
    CallByName obj, "ChangeBoolean", VbMethod, n
    Case36 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case37() As String
    Dim obj As New Item
    Dim n As Variant
    n = CBool(-1)
    On Error Resume Next
    CallByName obj, "ChangeBoolean", VbMethod, n
    Case37 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case38() As String
    Dim obj As New Item
    Dim n As Boolean
    n = CBool(-1)
    On Error Resume Next
    CallByName obj, "ChangeBoolean", VbMethod, (n)
    Case38 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case39() As String
    Dim obj As New Item
    Dim n As String
    n = "40"
    On Error Resume Next
    CallByName obj, "ChangeBoolean", VbMethod, n
    Case39 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case40() As String
    Dim obj As New Item
    Dim n As String
    n = CStr("abc")
    On Error Resume Next
    CallByName obj, "ChangeString", VbMethod, n
    Case40 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case41() As String
    Dim obj As New Item
    Dim n As Variant
    n = CStr("abc")
    On Error Resume Next
    CallByName obj, "ChangeString", VbMethod, n
    Case41 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case42() As String
    Dim obj As New Item
    Dim n As String
    n = CStr("abc")
    On Error Resume Next
    CallByName obj, "ChangeString", VbMethod, (n)
    Case42 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case43() As String
    Dim obj As New Item
    Dim n As Long
    n = CLng(40)
    On Error Resume Next
    CallByName obj, "ChangeString", VbMethod, n
    Case43 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Public Function Probe() As String
    Dim report As String
    report = report & "Byte_typed=" & Case0() & vbLf
    report = report & "Byte_variant=" & Case1() & vbLf
    report = report & "Byte_parentheses=" & Case2() & vbLf
    report = report & "Byte_wrong_type=" & Case3() & vbLf
    report = report & "Integer_typed=" & Case4() & vbLf
    report = report & "Integer_variant=" & Case5() & vbLf
    report = report & "Integer_parentheses=" & Case6() & vbLf
    report = report & "Integer_wrong_type=" & Case7() & vbLf
    report = report & "Long_typed=" & Case8() & vbLf
    report = report & "Long_variant=" & Case9() & vbLf
    report = report & "Long_parentheses=" & Case10() & vbLf
    report = report & "Long_wrong_type=" & Case11() & vbLf
    report = report & "LongLong_typed=" & Case12() & vbLf
    report = report & "LongLong_variant=" & Case13() & vbLf
    report = report & "LongLong_parentheses=" & Case14() & vbLf
    report = report & "LongLong_wrong_type=" & Case15() & vbLf
    report = report & "LongPtr_typed=" & Case16() & vbLf
    report = report & "LongPtr_variant=" & Case17() & vbLf
    report = report & "LongPtr_parentheses=" & Case18() & vbLf
    report = report & "LongPtr_wrong_type=" & Case19() & vbLf
    report = report & "Single_typed=" & Case20() & vbLf
    report = report & "Single_variant=" & Case21() & vbLf
    report = report & "Single_parentheses=" & Case22() & vbLf
    report = report & "Single_wrong_type=" & Case23() & vbLf
    report = report & "Double_typed=" & Case24() & vbLf
    report = report & "Double_variant=" & Case25() & vbLf
    report = report & "Double_parentheses=" & Case26() & vbLf
    report = report & "Double_wrong_type=" & Case27() & vbLf
    report = report & "Currency_typed=" & Case28() & vbLf
    report = report & "Currency_variant=" & Case29() & vbLf
    report = report & "Currency_parentheses=" & Case30() & vbLf
    report = report & "Currency_wrong_type=" & Case31() & vbLf
    report = report & "Date_typed=" & Case32() & vbLf
    report = report & "Date_variant=" & Case33() & vbLf
    report = report & "Date_parentheses=" & Case34() & vbLf
    report = report & "Date_wrong_type=" & Case35() & vbLf
    report = report & "Boolean_typed=" & Case36() & vbLf
    report = report & "Boolean_variant=" & Case37() & vbLf
    report = report & "Boolean_parentheses=" & Case38() & vbLf
    report = report & "Boolean_wrong_type=" & Case39() & vbLf
    report = report & "String_typed=" & Case40() & vbLf
    report = report & "String_variant=" & Case41() & vbLf
    report = report & "String_parentheses=" & Case42() & vbLf
    report = report & "String_wrong_type=" & Case43() & vbLf
    Probe = report
End Function
