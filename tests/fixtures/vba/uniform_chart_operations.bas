VERSION 1.0 CLASS
Attribute VB_Name = "ArchiveHost"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


VERSION 1.0 CLASS
Attribute VB_Name = "Ledger"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


Attribute VB_Name = "DriverModule"
' SPDX-License-Identifier: MPL-2.0
Option Explicit

Private Function Describe(ByVal target As Object) As String
    Dim title As String, position As String
    title = "[off]"
    position = "[off]"
    If target.HasTitle Then title = target.ChartTitle.Text
    If target.HasLegend Then position = CStr(target.Legend.Position)
    Describe = CStr(VarType(target)) & ":" & TypeName(target)
    Describe = Describe & "|title=" & CStr(target.HasTitle) & ":" & title
    Describe = Describe & "|legend=" & CStr(target.HasLegend) & ":" & position
    Describe = Describe & "|type=" & CStr(target.ChartType)
End Function

Public Function StateReport() As String
    Dim index As Long, anchor As Object, report As String
    On Error GoTo Failed
    report = "sheet=" & Describe(PlotHost) & "|data=" & DescribeSeries(PlotHost)
    report = report & "|embedded=" & CStr(Ledger.ChartObjects.Count)
    For index = 1 To Ledger.ChartObjects.Count
        Set anchor = Ledger.ChartObjects(index)
        report = report & "|object=" & anchor.Name & ":" & Describe(anchor.Chart) & "|data=" & DescribeSeries(anchor.Chart)
        report = report & "|bounds=" & CStr(anchor.Left) & ":" & CStr(anchor.Top)
        report = report & ":" & CStr(anchor.Width) & ":" & CStr(anchor.Height)
    Next index
    StateReport = report
    Exit Function
Failed:
    StateReport = "error|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function CaseProbe(ByVal kind As Long, ByVal useEmbedded As Boolean) As String
    Dim target As Object, anchor As Object
    On Error GoTo Failed
    If kind = 0 Then
        CaseProbe = "oxi-uniform-chart-20261106-v2"
        Exit Function
    End If
    If useEmbedded Then
        Set anchor = Ledger.ChartObjects.Add(20, 30, 240, 160)
        anchor.Name = "UniformEmbedded"
        Set target = anchor.Chart
    Else
        Set target = PlotHost
    End If
    Select Case kind
    Case 1
        ' Observe the original chart or the newly created embedded chart.
    Case 2
        target.HasTitle = True
    Case 3
        target.HasTitle = True
        target.ChartTitle.Text = "Revenue"
        target.HasLegend = False
    Case 4
        target.HasTitle = True
        target.ChartTitle.Text = ""
    Case 5
        target.HasTitle = False
        target.HasTitle = True
    Case 6
        target.HasTitle = False
    Case 7
        target.HasLegend = True
        target.Legend.Position = -4131
    Case 8
        target.HasLegend = True
        target.Legend.Position = -4107
    Case 9
        target.HasLegend = False
        target.HasLegend = True
    Case 10
        target.HasTitle = True
        target.ChartTitle.Text = ChrW(&H58F2) & ChrW(&H4E0A) & " & <" & ChrW(&H96C6) & ChrW(&H8A08) & ">" & vbLf & "2026"
    Case 11
        target.HasTitle = True
        target.ChartTitle.Caption = "Caption"
    Case 12
        target.ChartType = 51
    Case Else
        Err.Raise 5, "UniformChart", "Unknown case"
    End Select
    CaseProbe = "ok|" & Describe(target) & "|objects=" & CStr(Ledger.ChartObjects.Count)
    Exit Function
Failed:
    CaseProbe = "error|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Private Function TypedValues(ByVal value As Variant) As String
    Dim index As Long, result As String
    result = CStr(VarType(value)) & ":" & TypeName(value)
    If IsArray(value) Then
        result = result & "[" & CStr(LBound(value)) & "," & CStr(UBound(value)) & "]"
        For index = LBound(value) To UBound(value)
            result = result & "{" & CStr(VarType(value(index))) & ":" & CStr(value(index)) & "}"
        Next index
    Else
        result = result & ":" & CStr(value)
    End If
    TypedValues = result
End Function

Private Function DescribeSeries(ByVal target As Object) As String
    Dim index As Long, item As Object, result As String
    result = CStr(target.SeriesCollection.Count)
    For index = 1 To target.SeriesCollection.Count
        Set item = target.SeriesCollection(index)
        result = result & "|series=" & item.Name & ":" & item.Formula
        result = result & "|values=" & TypedValues(item.Values) & "|xs=" & TypedValues(item.XValues)
    Next index
    DescribeSeries = result
End Function

Public Function ChartTypeProbe(ByVal chartType As Long, ByVal sourceKind As Long, ByVal useEmbedded As Boolean) As String
    Dim target As Object, anchor As Object, item As Object
    On Error GoTo Failed
    If useEmbedded Then
        Set anchor = Ledger.ChartObjects.Add(20, 30, 240, 160)
        anchor.Name = "UniformEmbedded"
        Set target = anchor.Chart
    Else
        Set target = PlotHost
    End If
    target.ChartType = chartType
    If sourceKind = 1 Then
        Set item = target.SeriesCollection.NewSeries
        item.Name = "Revenue"
        item.Values = Array(1.25, 2.5, 3.75)
        item.XValues = Array("1", "2", "3")
    ElseIf sourceKind = 2 Then
        Ledger.Range("A1").Value2 = "Category"
        Ledger.Range("B1").Value2 = "Revenue"
        Ledger.Range("A2").Value2 = 1
        Ledger.Range("A3").Value2 = 2
        Ledger.Range("A4").Value2 = 3
        Ledger.Range("B2").Value2 = 1.25
        Ledger.Range("B3").Value2 = 2.5
        Ledger.Range("B4").Value2 = 3.75
        target.SetSourceData Ledger.Range("A1:B4")
    End If
    ChartTypeProbe = "ok|" & Describe(target) & "|data=" & DescribeSeries(target)
    Exit Function
Failed:
    ChartTypeProbe = "error|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

VERSION 1.0 CLASS
Attribute VB_Name = "PlotHost"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
