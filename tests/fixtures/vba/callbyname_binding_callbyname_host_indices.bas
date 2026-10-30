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

Attribute VB_Name = "Driver"
Option Explicit
Private Function DoCase(ByVal target As String, ByVal member As String, ByVal flag As Variant, ByVal scenario As Long) As String
    Dim obj As Object
    Dim result As Variant
    Dim savedNumber As Long
    ThisWorkbook.Worksheets(1).Name = "Sheet1"
    Range("A1").Value = ""
    Range("B1").Value = "payload"
    Select Case target
    Case "Range"
        Set obj = Range("A1")
    Case "Worksheet"
        Set obj = ThisWorkbook.Worksheets(1)
    Case "Workbook"
        Set obj = ThisWorkbook
    Case "Collection"
        Set obj = New Collection
        obj.Add "one"
    Case "Err"
        Set obj = Err
    End Select
    Err.Clear
    On Error Resume Next
    Select Case scenario
    Case 0
        result = CallByName(obj, member, flag)
    Case 1
        result = CallByName(obj, member, flag, "x")
    Case 2
        result = CallByName(obj, member, flag, Range("B1"))
    Case 3
        result = CallByName(obj, member, flag, 1, "x")
    Case 4
        result = CallByName(obj, member, flag, 10, "x")
    Case 5
        result = CallByName(obj, member, flag, 11, "x")
    Case 6
        result = CallByName(obj, member, flag, 12, "x")
    Case 7
        result = CallByName(obj, member, flag, 10, 1, "x")
    Case 8
        result = CallByName(obj, member, flag, 2, 1, "x")
    Case 9
        result = CallByName(obj, member, flag, 1)
    Case 10
        result = CallByName(obj, member, flag, 10)
    Case 11
        result = CallByName(obj, member, flag, Null)
    Case 12
        result = CallByName(obj, member, flag, Null, "x")
    Case 13
        result = CallByName(obj, member, flag, 9.6, "x")
    Case 14
        result = CallByName(obj, member, flag, "10", "x")
    End Select
    savedNumber = Err.Number
    DoCase = CStr(savedNumber) & "~" & TypeName(result) & "~" & CStr(Range("A1").Value)
End Function
Public Function Probe() As String
    Dim result As String
    result = result & "Value_2_9=" & DoCase("Range", "Value", 2, 9) & vbLf
    result = result & "Value_2_10=" & DoCase("Range", "Value", 2, 10) & vbLf
    result = result & "Value_2_11=" & DoCase("Range", "Value", 2, 11) & vbLf
    result = result & "Value_4_4=" & DoCase("Range", "Value", 4, 4) & vbLf
    result = result & "Value_4_5=" & DoCase("Range", "Value", 4, 5) & vbLf
    result = result & "Value_4_6=" & DoCase("Range", "Value", 4, 6) & vbLf
    result = result & "Value_4_7=" & DoCase("Range", "Value", 4, 7) & vbLf
    result = result & "Value_4_12=" & DoCase("Range", "Value", 4, 12) & vbLf
    result = result & "Value_4_13=" & DoCase("Range", "Value", 4, 13) & vbLf
    result = result & "Value_4_14=" & DoCase("Range", "Value", 4, 14) & vbLf
    result = result & "Value2_2_10=" & DoCase("Range", "Value2", 2, 10) & vbLf
    result = result & "Value2_4_4=" & DoCase("Range", "Value2", 4, 4) & vbLf
    result = result & "Item_2_0=" & DoCase("Range", "Item", 2, 0) & vbLf
    result = result & "Item_2_9=" & DoCase("Range", "Item", 2, 9) & vbLf
    result = result & "Item_4_1=" & DoCase("Range", "Item", 4, 1) & vbLf
    result = result & "Item_4_3=" & DoCase("Range", "Item", 4, 3) & vbLf
    result = result & "Item_4_8=" & DoCase("Range", "Item", 4, 8) & vbLf
    result = result & "_Default_1_0=" & DoCase("Range", "_Default", 1, 0) & vbLf
    result = result & "_Default_2_0=" & DoCase("Range", "_Default", 2, 0) & vbLf
    result = result & "_Default_4_1=" & DoCase("Range", "_Default", 4, 1) & vbLf
    result = result & "_Default_4_3=" & DoCase("Range", "_Default", 4, 3) & vbLf
    result = result & "_Default_4_8=" & DoCase("Range", "_Default", 4, 8) & vbLf
    Probe = result
End Function
