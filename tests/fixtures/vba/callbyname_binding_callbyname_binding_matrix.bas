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
Public HeldValue As Variant
Public Trace As String
Public Function MethodCall() As String
    MethodCall = "method"
End Function
Public Property Get ReadOnlyProp() As Variant
    ReadOnlyProp = "get"
End Property
Public Property Let OnlyLetProp(ByVal value As Variant)
    Trace = "let:" & TypeName(value)
End Property
Public Property Set OnlySetProp(ByVal value As Variant)
    Trace = "set:" & TypeName(value)
End Property
Public Property Let BothProp(ByVal value As Variant)
    Trace = "let:" & TypeName(value)
End Property
Public Property Set BothProp(ByVal value As Variant)
    Trace = "set:" & TypeName(value)
End Property
Public Property Get IndexedProp(ByVal index As Long) As Variant
    IndexedProp = "get:" & CStr(index)
End Property
Public Property Let IndexedProp(ByVal index As Long, ByVal value As Variant)
    Trace = "index:" & CStr(index) & ":" & TypeName(value)
End Property


Attribute VB_Name = "Driver"
Option Explicit
Private Function DoCase(ByVal member As String, ByVal flag As Variant, ByVal scenario As Long) As String
    Dim obj As New Item
    Dim result As Variant
    On Error Resume Next
    Select Case scenario
    Case 0
        result = CallByName(obj, member, flag)
    Case 1
        result = CallByName(obj, member, flag, "x")
    Case 2
        result = CallByName(obj, member, flag, Range("A1"))
    Case 3
        result = CallByName(obj, member, flag, 2, "x")
    End Select
    DoCase = CStr(Err.Number) & "~" & TypeName(result) & "~" & obj.Trace & "~" & TypeName(obj.HeldValue)
End Function
Public Function Probe() As String
    Dim result As String
    result = result & "HeldValue_noargs_1=" & DoCase("HeldValue", 1, 0) & vbLf
    result = result & "HeldValue_noargs_2=" & DoCase("HeldValue", 2, 0) & vbLf
    result = result & "HeldValue_noargs_3=" & DoCase("HeldValue", 3, 0) & vbLf
    result = result & "HeldValue_noargs_4=" & DoCase("HeldValue", 4, 0) & vbLf
    result = result & "HeldValue_noargs_8=" & DoCase("HeldValue", 8, 0) & vbLf
    result = result & "MethodCall_noargs_1=" & DoCase("MethodCall", 1, 0) & vbLf
    result = result & "MethodCall_noargs_2=" & DoCase("MethodCall", 2, 0) & vbLf
    result = result & "MethodCall_noargs_3=" & DoCase("MethodCall", 3, 0) & vbLf
    result = result & "MethodCall_noargs_4=" & DoCase("MethodCall", 4, 0) & vbLf
    result = result & "MethodCall_noargs_8=" & DoCase("MethodCall", 8, 0) & vbLf
    result = result & "ReadOnlyProp_noargs_1=" & DoCase("ReadOnlyProp", 1, 0) & vbLf
    result = result & "ReadOnlyProp_noargs_2=" & DoCase("ReadOnlyProp", 2, 0) & vbLf
    result = result & "ReadOnlyProp_noargs_3=" & DoCase("ReadOnlyProp", 3, 0) & vbLf
    result = result & "ReadOnlyProp_noargs_4=" & DoCase("ReadOnlyProp", 4, 0) & vbLf
    result = result & "ReadOnlyProp_noargs_8=" & DoCase("ReadOnlyProp", 8, 0) & vbLf
    result = result & "OnlyLetProp_noargs_1=" & DoCase("OnlyLetProp", 1, 0) & vbLf
    result = result & "OnlyLetProp_noargs_2=" & DoCase("OnlyLetProp", 2, 0) & vbLf
    result = result & "OnlyLetProp_noargs_3=" & DoCase("OnlyLetProp", 3, 0) & vbLf
    result = result & "OnlyLetProp_noargs_4=" & DoCase("OnlyLetProp", 4, 0) & vbLf
    result = result & "OnlyLetProp_noargs_8=" & DoCase("OnlyLetProp", 8, 0) & vbLf
    result = result & "OnlySetProp_noargs_1=" & DoCase("OnlySetProp", 1, 0) & vbLf
    result = result & "OnlySetProp_noargs_2=" & DoCase("OnlySetProp", 2, 0) & vbLf
    result = result & "OnlySetProp_noargs_3=" & DoCase("OnlySetProp", 3, 0) & vbLf
    result = result & "OnlySetProp_noargs_4=" & DoCase("OnlySetProp", 4, 0) & vbLf
    result = result & "OnlySetProp_noargs_8=" & DoCase("OnlySetProp", 8, 0) & vbLf
    result = result & "BothProp_noargs_1=" & DoCase("BothProp", 1, 0) & vbLf
    result = result & "BothProp_noargs_2=" & DoCase("BothProp", 2, 0) & vbLf
    result = result & "BothProp_noargs_3=" & DoCase("BothProp", 3, 0) & vbLf
    result = result & "BothProp_noargs_4=" & DoCase("BothProp", 4, 0) & vbLf
    result = result & "BothProp_noargs_8=" & DoCase("BothProp", 8, 0) & vbLf
    result = result & "IndexedProp_noargs_1=" & DoCase("IndexedProp", 1, 0) & vbLf
    result = result & "IndexedProp_noargs_2=" & DoCase("IndexedProp", 2, 0) & vbLf
    result = result & "IndexedProp_noargs_3=" & DoCase("IndexedProp", 3, 0) & vbLf
    result = result & "IndexedProp_noargs_4=" & DoCase("IndexedProp", 4, 0) & vbLf
    result = result & "IndexedProp_noargs_8=" & DoCase("IndexedProp", 8, 0) & vbLf
    result = result & "HeldValue_4_1=" & DoCase("HeldValue", 4, 1) & vbLf
    result = result & "HeldValue_4_2=" & DoCase("HeldValue", 4, 2) & vbLf
    result = result & "HeldValue_8_1=" & DoCase("HeldValue", 8, 1) & vbLf
    result = result & "HeldValue_8_2=" & DoCase("HeldValue", 8, 2) & vbLf
    result = result & "OnlyLetProp_4_1=" & DoCase("OnlyLetProp", 4, 1) & vbLf
    result = result & "OnlyLetProp_4_2=" & DoCase("OnlyLetProp", 4, 2) & vbLf
    result = result & "OnlyLetProp_8_1=" & DoCase("OnlyLetProp", 8, 1) & vbLf
    result = result & "OnlyLetProp_8_2=" & DoCase("OnlyLetProp", 8, 2) & vbLf
    result = result & "OnlySetProp_4_1=" & DoCase("OnlySetProp", 4, 1) & vbLf
    result = result & "OnlySetProp_4_2=" & DoCase("OnlySetProp", 4, 2) & vbLf
    result = result & "OnlySetProp_8_1=" & DoCase("OnlySetProp", 8, 1) & vbLf
    result = result & "OnlySetProp_8_2=" & DoCase("OnlySetProp", 8, 2) & vbLf
    result = result & "BothProp_4_1=" & DoCase("BothProp", 4, 1) & vbLf
    result = result & "BothProp_4_2=" & DoCase("BothProp", 4, 2) & vbLf
    result = result & "BothProp_8_1=" & DoCase("BothProp", 8, 1) & vbLf
    result = result & "BothProp_8_2=" & DoCase("BothProp", 8, 2) & vbLf
    result = result & "IndexedProp_indexed_1=" & DoCase("IndexedProp", 1, 3) & vbLf
    result = result & "IndexedProp_indexed_2=" & DoCase("IndexedProp", 2, 3) & vbLf
    result = result & "IndexedProp_indexed_3=" & DoCase("IndexedProp", 3, 3) & vbLf
    result = result & "IndexedProp_indexed_4=" & DoCase("IndexedProp", 4, 3) & vbLf
    result = result & "IndexedProp_indexed_8=" & DoCase("IndexedProp", 8, 3) & vbLf
    result = result & "flag_coercion_1.5=" & DoCase("ReadOnlyProp", 1.5, 0) & vbLf
    result = result & "flag_coercion_2.5=" & DoCase("ReadOnlyProp", 2.5, 0) & vbLf
    result = result & "flag_coercion_0=" & DoCase("ReadOnlyProp", 0, 0) & vbLf
    result = result & "flag_coercion_-1=" & DoCase("ReadOnlyProp", -1, 0) & vbLf
    result = result & "flag_coercion_16=" & DoCase("ReadOnlyProp", 16, 0) & vbLf
    result = result & "flag_coercion_Null=" & DoCase("ReadOnlyProp", Null, 0) & vbLf
    result = result & "flag_coercion_""2""=" & DoCase("ReadOnlyProp", "2", 0) & vbLf
    result = result & "flag_coercion_""bad""=" & DoCase("ReadOnlyProp", "bad", 0) & vbLf
    Probe = result
End Function
