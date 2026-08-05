# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-string-statements-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$excel = $null
$workbook = $null
$sheet = $null
$component = $null

$source = @'
Option Explicit

Public Sub OxiStringStatementsProbe()
    Dim value As String

    value = "abcdef"
    Mid$(value, 2, 3) = "WXYZ"
    ThisWorkbook.Worksheets(1).Cells(1, 1).Value2 = value

    value = "abcdef"
    Mid(value, 2) = "Q"
    ThisWorkbook.Worksheets(1).Cells(2, 1).Value2 = value

    value = "abcdef"
    LSet value = "xy"
    ThisWorkbook.Worksheets(1).Cells(3, 1).Value2 = "[" & value & "]"
    ThisWorkbook.Worksheets(1).Cells(4, 1).Value2 = Len(value)

    value = "abcdef"
    RSet value = "xy"
    ThisWorkbook.Worksheets(1).Cells(5, 1).Value2 = "[" & value & "]"
    ThisWorkbook.Worksheets(1).Cells(6, 1).Value2 = Len(value)

    value = "abcdef"
    On Error Resume Next
    Mid$(value, 10, 1) = "Z"
    ThisWorkbook.Worksheets(1).Cells(7, 1).Value2 = value
    ThisWorkbook.Worksheets(1).Cells(8, 1).Value2 = Err.Number
    Err.Clear

    Mid$(value, 0, 1) = "Z"
    ThisWorkbook.Worksheets(1).Cells(9, 1).Value2 = Err.Number
    Err.Clear

    Mid$(value, 2, 0) = "Z"
    ThisWorkbook.Worksheets(1).Cells(10, 1).Value2 = Err.Number
    Err.Clear
    On Error GoTo 0
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

    $component.Name = 'OxiStringStatementsModule'
    $component.CodeModule.AddFromString($source)
    $excel.Run("'$($workbook.Name)'!OxiStringStatementsModule.OxiStringStatementsProbe")
    $sheet = $workbook.Worksheets.Item(1)

    $result = [ordered]@{
        mid_with_length = [string]$sheet.Cells.Item(1, 1).Value2
        mid_without_length = [string]$sheet.Cells.Item(2, 1).Value2
        lset_value = [string]$sheet.Cells.Item(3, 1).Value2
        lset_length = [long]$sheet.Cells.Item(4, 1).Value2
        rset_value = [string]$sheet.Cells.Item(5, 1).Value2
        rset_length = [long]$sheet.Cells.Item(6, 1).Value2
        start_past_end_value = [string]$sheet.Cells.Item(7, 1).Value2
        start_past_end_error = [long]$sheet.Cells.Item(8, 1).Value2
        start_zero_error = [long]$sheet.Cells.Item(9, 1).Value2
        zero_length_error = [long]$sheet.Cells.Item(10, 1).Value2
    }
    $expected = [ordered]@{
        mid_with_length = 'aWXYef'
        mid_without_length = 'aQcdef'
        lset_value = '[xy    ]'
        lset_length = 6L
        rset_value = '[    xy]'
        rset_length = 6L
        start_past_end_value = 'abcdef'
        start_past_end_error = 5L
        start_zero_error = 5L
        zero_length_error = 0L
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
