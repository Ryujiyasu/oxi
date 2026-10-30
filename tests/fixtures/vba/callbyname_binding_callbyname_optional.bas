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
Public Property Let Choice(Optional ByVal index As Long = 7, ByVal value As Variant)
    Trace = "let:" & CStr(index) & ":" & TypeName(value)
End Property
Public Property Set Choice(Optional ByVal index As Long = 7, ByVal value As Variant)
    Trace = "set:" & CStr(index) & ":" & TypeName(value)
End Property
Public Property Get Choice(Optional ByVal index As Long = 7) As Variant
    Choice = CStr(index)
End Property
Public Sub Touch(ByRef value As Long)
    value = value + 2
End Sub


Attribute VB_Name = "Driver"
Option Explicit
Public Function Probe() As String
    Dim obj As New Item
    Dim result As Variant
    Dim report As String
    Dim n As Long
    Dim dict As Object
    On Error Resume Next
    CallByName obj, "Choice", VbLet, "x"
    report = "optional_let=" & CStr(Err.Number) & "~" & obj.Trace & vbLf
    Err.Clear
    CallByName obj, "Choice", VbSet, "x"
    report = report & "optional_set=" & CStr(Err.Number) & "~" & obj.Trace & vbLf
    Err.Clear
    CallByName obj, "Choice", VbLet, 2, "x"
    report = report & "explicit_let=" & CStr(Err.Number) & "~" & obj.Trace & vbLf
    Err.Clear
    n = 40
    CallByName obj, "Touch", VbMethod, n
    report = report & "byref_copyback=" & CStr(Err.Number) & "~" & CStr(n) & vbLf
    Err.Clear
    Range("C1").value = "Choice"
    Range("D1").value = 2.6
    result = CallByName(obj, Range("C1"), Range("D1"))
    report = report & "object_name_kind=" & CStr(Err.Number) & "~" & CStr(result) & vbLf
    Err.Clear
    Set dict = CreateObject("Scripting.Dictionary")
    Range("A1").value = "payload"
    CallByName dict, "Item", VbLet, "one", Range("A1")
    report = report & "dictionary_object=" & CStr(Err.Number) & "~" & TypeName(dict.Item("one")) & vbLf
    Probe = report
End Function

