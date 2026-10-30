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

Public Function NeedArg(ByVal value As Long) As String
    NeedArg = CStr(value)
End Function


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
    result = result & "NeedArg_1_0=" & DoCase("NeedArg", 1, 0) & vbLf
    result = result & "NeedArg_2_0=" & DoCase("NeedArg", 2, 0) & vbLf
    result = result & "NeedArg_3_0=" & DoCase("NeedArg", 3, 0) & vbLf
    result = result & "IndexedProp_4_1=" & DoCase("IndexedProp", 4, 1) & vbLf
    result = result & "IndexedProp_4_2=" & DoCase("IndexedProp", 4, 2) & vbLf
    result = result & "HeldValue_2_1=" & DoCase("HeldValue", 2, 1) & vbLf
    result = result & "HeldValue_3_1=" & DoCase("HeldValue", 3, 1) & vbLf
    result = result & "HeldValue_4_3=" & DoCase("HeldValue", 4, 3) & vbLf
    result = result & "HeldValue_8_3=" & DoCase("HeldValue", 8, 3) & vbLf
    result = result & "OnlyLetProp_4_3=" & DoCase("OnlyLetProp", 4, 3) & vbLf
    Probe = result
End Function
