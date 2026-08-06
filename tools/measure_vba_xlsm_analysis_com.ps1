# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-xlsm-analysis-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$workbookCopyPath = Join-Path $probeRoot 'probe-copy.xlsm'
$workbookVariantPath = Join-Path $probeRoot 'probe-variant.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$module = $null
$sheet = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'AnalysisProbe'
    $module.CodeModule.AddFromString(@'
Option Explicit

Public Sub BuildReport(ByVal target As Worksheet)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = SuffixValue$(42&)
    ws.Range("B1").Value = Len(SuffixValue$(42&))
    ws.Range("A1").Font.Bold = True
End Sub

Private Function HiddenHelper(ByVal value As Long) As Long
    HiddenHelper = value + 1
End Function

Private Function SuffixValue$(ByVal number&)
    Dim padded As String * 4
    Dim chunks() As String
    Dim lazy As New Collection
    ReDim chunks(1 To 2) As String
    chunks(1) = CStr(number&)
    lazy.Add chunks(1)
    padded = lazy(1)
    SuffixValue$ = padded
End Function
'@)
    $sheet = $workbook.Worksheets.Item(1)
    $excel.Run("'$($workbook.Name)'!AnalysisProbe.BuildReport", $sheet)
    $actualValue = [string]$sheet.Range('A1').Value2
    $fixedLength = [long]$sheet.Range('B1').Value2
    if ($actualValue -ne '42' -or $fixedLength -ne 4 -or -not [bool]$sheet.Range('A1').Font.Bold) {
        throw "COM execution mismatch: value=[$actualValue], fixed-length=$fixedLength, bold=$([bool]$sheet.Range('A1').Font.Bold)"
    }
    $workbook.SaveAs($workbookPath, 52)
    $workbook.SaveCopyAs($workbookCopyPath)
    $module.CodeModule.ReplaceLine(12, '    HiddenHelper = value + 2')
    $module.CodeModule.InsertLines(2, 'Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr')
    $module.CodeModule.AddFromString(@'

Private Function VariantOnly(ByVal value As Long) As Long
    VariantOnly = SetTimer(0, 0, value, AddressOf HiddenHelper)
End Function
'@)
    $workbook.SaveAs($workbookVariantPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    & cargo build -q -p oxidocs-cli --manifest-path (Join-Path $repoRoot 'Cargo.toml')
    if ($LASTEXITCODE -ne 0) {
        throw "oxidocs-cli build failed with exit code $LASTEXITCODE"
    }

    $output = (& $cliPath vba-analyze $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-analyze failed with exit code $LASTEXITCODE"
    }

    $expectations = @(
        '[AnalysisProbe]',
        'verdict: A (report generation)',
        'procedures: 3, statements: 14',
        'unparsed: 0',
        'Summary:'
    )
    foreach ($expected in $expectations) {
        if (-not $output.Contains($expected)) {
            throw "Expected output fragment not found: $expected`n$output"
        }
    }

    $output.TrimEnd()
    $inventory = (& $cliPath vba-inventory $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory failed with exit code $LASTEXITCODE"
    }
    $inventoryExpectations = @(
        'Structurally identical modules (standard fingerprint):',
        'probe.xlsm::AnalysisProbe',
        'probe-copy.xlsm::AnalysisProbe',
        'Related modules (standard fingerprint):',
        'probe-variant.xlsm::AnalysisProbe',
        '40.0% (shared 2; only 1/2; diverged: HiddenHelper; declarations differ)',
        'Inventory: 3 succeeded, 0 failed'
    )
    foreach ($expected in $inventoryExpectations) {
        if (-not $inventory.Contains($expected)) {
            throw "Expected inventory fragment not found: $expected`n$inventory"
        }
    }
    $inventory.TrimEnd()

    $jsonOutput = (& $cliPath vba-inventory-json $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    if ($jsonReport.schema -ne 'oxivba-inventory-v1') {
        throw "Unexpected JSON schema: $($jsonReport.schema)"
    }
    if (@($jsonReport.projects).Count -ne 3) {
        throw "Expected 3 JSON projects, got $(@($jsonReport.projects).Count)"
    }
    if (@($jsonReport.duplicate_groups).Count -ne 1) {
        throw "Expected 1 JSON duplicate group, got $(@($jsonReport.duplicate_groups).Count)"
    }
    if (@($jsonReport.related_modules).Count -ne 2) {
        throw "Expected 2 JSON related pairs, got $(@($jsonReport.related_modules).Count)"
    }
    $divergedNames = @($jsonReport.related_modules | ForEach-Object { $_.diverged })
    if ($divergedNames -notcontains 'HiddenHelper') {
        throw "JSON related pairs did not report HiddenHelper divergence"
    }
    if (@($jsonReport.related_modules | Where-Object { $_.declarations_differ }).Count -ne 2) {
        throw "JSON related pairs did not report the changed Declare context"
    }
    $variantProject = @($jsonReport.projects | Where-Object { $_.path -like '*probe-variant.xlsm' })[0]
    $variantModule = @($variantProject.modules | Where-Object { $_.name -eq 'AnalysisProbe' })[0]
    if ($variantModule.metrics.unparsed -ne 0) {
        throw "AddressOf variant produced unparsed source"
    }
    if (@($variantModule.external_declares) -notcontains 'SetTimer') {
        throw "AddressOf variant did not report its SetTimer declaration"
    }
    if ($variantModule.api_names.Worksheet -ne 2) {
        throw "Procedure and local declaration types were not both reported"
    }
    if ($variantModule.api_names.Collection -ne 1) {
        throw "As New declaration type was not reported"
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw "JSON inventory reported unexpected errors"
    }
    Write-Output 'JSON inventory schema round-trip: PASS'
    Write-Output 'COM XLSM extraction and analysis probe: PASS'
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($comObject in @($sheet, $module, $workbook, $excel)) {
        if ($null -ne $comObject) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $probeRoot) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}
