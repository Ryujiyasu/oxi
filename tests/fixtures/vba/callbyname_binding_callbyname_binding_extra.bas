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

Private Function HiddenPrivate() As String
    HiddenPrivate = "private"
End Function
Friend Function HiddenFriend() As String
    HiddenFriend = "friend"
End Function


Attribute VB_Name = "Driver"
Option Explicit
Private Function DoCase(ByVal member As Variant, ByVal flag As Variant, ByVal scenario As Long) As String
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
    result = result & "NoSuchMember_noargs_1=" & DoCase("NoSuchMember", 1, 0) & vbLf
    result = result & "NoSuchMember_noargs_2=" & DoCase("NoSuchMember", 2, 0) & vbLf
    result = result & "NoSuchMember_noargs_3=" & DoCase("NoSuchMember", 3, 0) & vbLf
    result = result & "NoSuchMember_noargs_4=" & DoCase("NoSuchMember", 4, 0) & vbLf
    result = result & "NoSuchMember_noargs_8=" & DoCase("NoSuchMember", 8, 0) & vbLf
    result = result & "HiddenPrivate_noargs_1=" & DoCase("HiddenPrivate", 1, 0) & vbLf
    result = result & "HiddenPrivate_noargs_2=" & DoCase("HiddenPrivate", 2, 0) & vbLf
    result = result & "HiddenPrivate_noargs_3=" & DoCase("HiddenPrivate", 3, 0) & vbLf
    result = result & "HiddenPrivate_noargs_4=" & DoCase("HiddenPrivate", 4, 0) & vbLf
    result = result & "HiddenPrivate_noargs_8=" & DoCase("HiddenPrivate", 8, 0) & vbLf
    result = result & "HiddenFriend_noargs_1=" & DoCase("HiddenFriend", 1, 0) & vbLf
    result = result & "HiddenFriend_noargs_2=" & DoCase("HiddenFriend", 2, 0) & vbLf
    result = result & "HiddenFriend_noargs_3=" & DoCase("HiddenFriend", 3, 0) & vbLf
    result = result & "HiddenFriend_noargs_4=" & DoCase("HiddenFriend", 4, 0) & vbLf
    result = result & "HiddenFriend_noargs_8=" & DoCase("HiddenFriend", 8, 0) & vbLf
    result = result & "unknown_put_4_1=" & DoCase("NoSuchMember", 4, 1) & vbLf
    result = result & "unknown_put_4_2=" & DoCase("NoSuchMember", 4, 2) & vbLf
    result = result & "unknown_put_8_1=" & DoCase("NoSuchMember", 8, 1) & vbLf
    result = result & "unknown_put_8_2=" & DoCase("NoSuchMember", 8, 2) & vbLf
    result = result & "null_name_1=" & DoCase(Null, 1, 0) & vbLf
    result = result & "null_name_2=" & DoCase(Null, 2, 0) & vbLf
    result = result & "null_name_4=" & DoCase(Null, 4, 0) & vbLf
    result = result & "null_name_8=" & DoCase(Null, 8, 0) & vbLf
    result = result & "empty_name_1=" & DoCase("", 1, 0) & vbLf
    result = result & "empty_name_2=" & DoCase("", 2, 0) & vbLf
    result = result & "empty_name_4=" & DoCase("", 4, 0) & vbLf
    result = result & "empty_name_8=" & DoCase("", 8, 0) & vbLf
    result = result & "rounding_1.6=" & DoCase("MethodCall", 1.6, 0) & vbLf
    result = result & "rounding_2.6=" & DoCase("MethodCall", 2.6, 0) & vbLf
    result = result & "rounding_7.5=" & DoCase("MethodCall", 7.5, 0) & vbLf
    result = result & "rounding_2147483648#=" & DoCase("MethodCall", 2147483648#, 0) & vbLf
    Probe = result
End Function
