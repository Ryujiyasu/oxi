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
    End Select
    savedNumber = Err.Number
    DoCase = CStr(savedNumber) & "~" & TypeName(result) & "~" & CStr(Range("A1").Value)
End Function
Public Function Probe() As String
    Dim result As String
    result = result & "Range_Value_noargs_1=" & DoCase("Range", "Value", 1, 0) & vbLf
    result = result & "Range_Value_noargs_2=" & DoCase("Range", "Value", 2, 0) & vbLf
    result = result & "Range_Value_noargs_3=" & DoCase("Range", "Value", 3, 0) & vbLf
    result = result & "Range_Value_noargs_4=" & DoCase("Range", "Value", 4, 0) & vbLf
    result = result & "Range_Value_noargs_8=" & DoCase("Range", "Value", 8, 0) & vbLf
    result = result & "Range_Address_noargs_1=" & DoCase("Range", "Address", 1, 0) & vbLf
    result = result & "Range_Address_noargs_2=" & DoCase("Range", "Address", 2, 0) & vbLf
    result = result & "Range_Address_noargs_3=" & DoCase("Range", "Address", 3, 0) & vbLf
    result = result & "Range_Address_noargs_4=" & DoCase("Range", "Address", 4, 0) & vbLf
    result = result & "Range_Address_noargs_8=" & DoCase("Range", "Address", 8, 0) & vbLf
    result = result & "Range_Clear_noargs_1=" & DoCase("Range", "Clear", 1, 0) & vbLf
    result = result & "Range_Clear_noargs_2=" & DoCase("Range", "Clear", 2, 0) & vbLf
    result = result & "Range_Clear_noargs_3=" & DoCase("Range", "Clear", 3, 0) & vbLf
    result = result & "Range_Clear_noargs_4=" & DoCase("Range", "Clear", 4, 0) & vbLf
    result = result & "Range_Clear_noargs_8=" & DoCase("Range", "Clear", 8, 0) & vbLf
    result = result & "Worksheet_Name_noargs_1=" & DoCase("Worksheet", "Name", 1, 0) & vbLf
    result = result & "Worksheet_Name_noargs_2=" & DoCase("Worksheet", "Name", 2, 0) & vbLf
    result = result & "Worksheet_Name_noargs_3=" & DoCase("Worksheet", "Name", 3, 0) & vbLf
    result = result & "Worksheet_Name_noargs_4=" & DoCase("Worksheet", "Name", 4, 0) & vbLf
    result = result & "Worksheet_Name_noargs_8=" & DoCase("Worksheet", "Name", 8, 0) & vbLf
    result = result & "Worksheet_Calculate_noargs_1=" & DoCase("Worksheet", "Calculate", 1, 0) & vbLf
    result = result & "Worksheet_Calculate_noargs_2=" & DoCase("Worksheet", "Calculate", 2, 0) & vbLf
    result = result & "Worksheet_Calculate_noargs_3=" & DoCase("Worksheet", "Calculate", 3, 0) & vbLf
    result = result & "Worksheet_Calculate_noargs_4=" & DoCase("Worksheet", "Calculate", 4, 0) & vbLf
    result = result & "Worksheet_Calculate_noargs_8=" & DoCase("Worksheet", "Calculate", 8, 0) & vbLf
    result = result & "Workbook_Name_noargs_1=" & DoCase("Workbook", "Name", 1, 0) & vbLf
    result = result & "Workbook_Name_noargs_2=" & DoCase("Workbook", "Name", 2, 0) & vbLf
    result = result & "Workbook_Name_noargs_3=" & DoCase("Workbook", "Name", 3, 0) & vbLf
    result = result & "Workbook_Name_noargs_4=" & DoCase("Workbook", "Name", 4, 0) & vbLf
    result = result & "Workbook_Name_noargs_8=" & DoCase("Workbook", "Name", 8, 0) & vbLf
    result = result & "Collection_Count_noargs_1=" & DoCase("Collection", "Count", 1, 0) & vbLf
    result = result & "Collection_Count_noargs_2=" & DoCase("Collection", "Count", 2, 0) & vbLf
    result = result & "Collection_Count_noargs_3=" & DoCase("Collection", "Count", 3, 0) & vbLf
    result = result & "Collection_Count_noargs_4=" & DoCase("Collection", "Count", 4, 0) & vbLf
    result = result & "Collection_Count_noargs_8=" & DoCase("Collection", "Count", 8, 0) & vbLf
    result = result & "Collection_Add_noargs_1=" & DoCase("Collection", "Add", 1, 0) & vbLf
    result = result & "Collection_Add_noargs_2=" & DoCase("Collection", "Add", 2, 0) & vbLf
    result = result & "Collection_Add_noargs_3=" & DoCase("Collection", "Add", 3, 0) & vbLf
    result = result & "Collection_Add_noargs_4=" & DoCase("Collection", "Add", 4, 0) & vbLf
    result = result & "Collection_Add_noargs_8=" & DoCase("Collection", "Add", 8, 0) & vbLf
    result = result & "Err_Number_noargs_1=" & DoCase("Err", "Number", 1, 0) & vbLf
    result = result & "Err_Number_noargs_2=" & DoCase("Err", "Number", 2, 0) & vbLf
    result = result & "Err_Number_noargs_3=" & DoCase("Err", "Number", 3, 0) & vbLf
    result = result & "Err_Number_noargs_4=" & DoCase("Err", "Number", 4, 0) & vbLf
    result = result & "Err_Number_noargs_8=" & DoCase("Err", "Number", 8, 0) & vbLf
    result = result & "Err_Clear_noargs_1=" & DoCase("Err", "Clear", 1, 0) & vbLf
    result = result & "Err_Clear_noargs_2=" & DoCase("Err", "Clear", 2, 0) & vbLf
    result = result & "Err_Clear_noargs_3=" & DoCase("Err", "Clear", 3, 0) & vbLf
    result = result & "Err_Clear_noargs_4=" & DoCase("Err", "Clear", 4, 0) & vbLf
    result = result & "Err_Clear_noargs_8=" & DoCase("Err", "Clear", 8, 0) & vbLf
    result = result & "Range_Value_4_1=" & DoCase("Range", "Value", 4, 1) & vbLf
    result = result & "Range_Value_4_2=" & DoCase("Range", "Value", 4, 2) & vbLf
    result = result & "Range_Value_4_3=" & DoCase("Range", "Value", 4, 3) & vbLf
    result = result & "Range_Value_8_1=" & DoCase("Range", "Value", 8, 1) & vbLf
    result = result & "Range_Value_8_2=" & DoCase("Range", "Value", 8, 2) & vbLf
    result = result & "Range_Value_8_3=" & DoCase("Range", "Value", 8, 3) & vbLf
    result = result & "Worksheet_Name_4_1=" & DoCase("Worksheet", "Name", 4, 1) & vbLf
    result = result & "Worksheet_Name_4_2=" & DoCase("Worksheet", "Name", 4, 2) & vbLf
    result = result & "Worksheet_Name_4_3=" & DoCase("Worksheet", "Name", 4, 3) & vbLf
    result = result & "Worksheet_Name_8_1=" & DoCase("Worksheet", "Name", 8, 1) & vbLf
    result = result & "Worksheet_Name_8_2=" & DoCase("Worksheet", "Name", 8, 2) & vbLf
    result = result & "Worksheet_Name_8_3=" & DoCase("Worksheet", "Name", 8, 3) & vbLf
    Probe = result
End Function
