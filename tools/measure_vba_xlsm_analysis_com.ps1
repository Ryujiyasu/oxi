# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-xlsm-analysis-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$workbookCopyPath = Join-Path $probeRoot 'probe-copy.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$module = $null

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

Public Sub BuildReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = 42
    ws.Range("A1").Font.Bold = True
End Sub

Private Function HiddenHelper(ByVal value As Long) As Long
    HiddenHelper = value + 1
End Function
'@)
    $workbook.SaveAs($workbookPath, 52)
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
        'procedures: 2, statements: 5',
        'unparsed: 0',
        'Summary:'
    )
    foreach ($expected in $expectations) {
        if (-not $output.Contains($expected)) {
            throw "Expected output fragment not found: $expected`n$output"
        }
    }

    $output.TrimEnd()
    Copy-Item -LiteralPath $workbookPath -Destination $workbookCopyPath
    $inventory = (& $cliPath vba-inventory $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory failed with exit code $LASTEXITCODE"
    }
    $inventoryExpectations = @(
        'Structurally identical modules (standard fingerprint):',
        'probe.xlsm::AnalysisProbe',
        'probe-copy.xlsm::AnalysisProbe',
        'Inventory: 2 succeeded, 0 failed'
    )
    foreach ($expected in $inventoryExpectations) {
        if (-not $inventory.Contains($expected)) {
            throw "Expected inventory fragment not found: $expected`n$inventory"
        }
    }
    $inventory.TrimEnd()
    Write-Output 'COM XLSM extraction and analysis probe: PASS'
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($comObject in @($module, $workbook, $excel)) {
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
