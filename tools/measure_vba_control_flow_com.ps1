# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-control-flow-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$excel = $null
$workbook = $null
$sheet = $null
$component = $null

$source = @'
Option Explicit

Private Function DispatchGoTo(ByVal choice As Long) As String
    DispatchGoTo = "fallthrough"
    On choice GoTo First, Second, Third
    Exit Function
First:
    DispatchGoTo = "first"
    Exit Function
Second:
    DispatchGoTo = "second"
    Exit Function
Third:
    DispatchGoTo = "third"
End Function

Private Function DispatchGoSub(ByVal choice As Long) As String
    DispatchGoSub = ""
    On choice GoSub First, Second, Third
AfterDispatch:
    DispatchGoSub = DispatchGoSub & "after"
    Exit Function
First:
    DispatchGoSub = "first:"
    Return
Second:
    DispatchGoSub = "second:"
    Return
Third:
    DispatchGoSub = "third:"
    Return
End Function

Private Function NegativeGoToError() As Long
    On Error Resume Next
    On -1 GoTo Target
    NegativeGoToError = Err.Number
    Exit Function
Target:
    NegativeGoToError = -1
End Function

Private Function NegativeGoSubError() As Long
    On Error Resume Next
    On -1 GoSub Target
    NegativeGoSubError = Err.Number
    Exit Function
Target:
    Return
End Function

Public Sub OxiControlFlowProbe()
    Dim choices As Variant
    Dim i As Long
    choices = Array(0, 1, 2, 3, 4)
    For i = LBound(choices) To UBound(choices)
        ThisWorkbook.Worksheets(1).Cells(i + 1, 1).Value2 = DispatchGoTo(choices(i))
        ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value2 = DispatchGoSub(choices(i))
    Next i
    ThisWorkbook.Worksheets(1).Cells(6, 1).Value2 = NegativeGoToError()
    ThisWorkbook.Worksheets(1).Cells(6, 2).Value2 = NegativeGoSubError()
End Sub
'@

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($workbookPath, 52)

    try {
        $component = $workbook.VBProject.VBComponents.Add(1)
    }
    catch {
        throw "Excel COM is available, but VBProject access is disabled. $($_.Exception.Message)"
    }
    if ($null -eq $component) {
        throw 'Excel returned no VBComponent. Enable trust access to the VBA project object model.'
    }

    $component.Name = 'OxiControlFlowModule'
    $component.CodeModule.AddFromString($source)
    $excel.Run("'$($workbook.Name)'!OxiControlFlowModule.OxiControlFlowProbe")
    $sheet = $workbook.Worksheets.Item(1)

    $result = [ordered]@{}
    foreach ($row in 1..5) {
        $choice = @(0, 1, 2, 3, 4)[$row - 1]
        $result["goto_${choice}"] = [string]$sheet.Cells.Item($row, 1).Value2
        $result["gosub_${choice}"] = [string]$sheet.Cells.Item($row, 2).Value2
    }
    $result['goto_negative_error'] = [long]$sheet.Cells.Item(6, 1).Value2
    $result['gosub_negative_error'] = [long]$sheet.Cells.Item(6, 2).Value2
    $expected = [ordered]@{
        'goto_0' = 'fallthrough'
        'goto_1' = 'first'
        'goto_2' = 'second'
        'goto_3' = 'third'
        'goto_4' = 'fallthrough'
        'gosub_0' = 'after'
        'gosub_1' = 'first:after'
        'gosub_2' = 'second:after'
        'gosub_3' = 'third:after'
        'gosub_4' = 'after'
        'goto_negative_error' = 5L
        'gosub_negative_error' = 5L
    }
    foreach ($key in $expected.Keys) {
        if ($result[$key] -ne $expected[$key]) {
            throw "COM result mismatch for ${key}: expected '$($expected[$key])', got '$($result[$key])'"
        }
    }
    $result | ConvertTo-Json
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($comObject in @($sheet, $component, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    if (Test-Path -LiteralPath $probeRoot) {
        [IO.Directory]::Delete($probeRoot, $true)
    }
}
