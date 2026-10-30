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
Public Sub BumpLong(ByRef value As Long)
    Trace = TypeName(value)
    value = value + 2
End Sub
Public Sub BumpLongLong(ByRef value As LongLong)
    Trace = TypeName(value)
    value = value + 2
End Sub
Public Sub BumpLongPtr(ByRef value As LongPtr)
    Trace = TypeName(value)
    value = value + 2
End Sub
Public Sub BumpEnum(ByRef value As Shade)
    Trace = TypeName(value)
    value = value + 2
End Sub


Attribute VB_Name = "Driver"
Option Explicit
Public Enum Shade
    Dark = 1
End Enum
Private Function Case0() As String
    Dim obj As New Item
    Dim n As Shade
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLong", VbMethod, n
    Case0 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case1() As String
    Dim obj As New Item
    Dim n As Long
    n = 40
    On Error Resume Next
    CallByName obj, "BumpEnum", VbMethod, n
    Case1 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case2() As String
    Dim obj As New Item
    Dim n As Shade
    n = 40
    On Error Resume Next
    CallByName obj, "BumpEnum", VbMethod, n
    Case2 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case3() As String
    Dim obj As New Item
    Dim n As LongPtr
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLongLong", VbMethod, n
    Case3 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case4() As String
    Dim obj As New Item
    Dim n As LongLong
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLongPtr", VbMethod, n
    Case4 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Private Function Case5() As String
    Dim obj As New Item
    Dim n As LongPtr
    n = 40
    On Error Resume Next
    CallByName obj, "BumpLongPtr", VbMethod, n
    Case5 = CStr(Err.Number) & "~" & CStr(n) & "~" & TypeName(n) & "~" & obj.Trace
End Function
Public Function Probe() As String
    Dim report As String
    report = report & "enum_to_long=" & Case0() & vbLf
    report = report & "long_to_enum=" & Case1() & vbLf
    report = report & "enum_to_enum=" & Case2() & vbLf
    report = report & "pointer_to_longlong=" & Case3() & vbLf
    report = report & "longlong_to_pointer=" & Case4() & vbLf
    report = report & "pointer_to_pointer=" & Case5() & vbLf
    Probe = report
End Function
