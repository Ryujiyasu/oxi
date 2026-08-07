# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-excel4-macro-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'excel4-macro-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'Excel4MacroProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function LegacySum() As Double
    LegacySum = Application.ExecuteExcel4Macro("SUM(40,2)")
End Function
'@)

    $actual = [double]$excel.Run("'$($workbook.Name)'!Excel4MacroProbe.LegacySum")
    if ($actual -ne 42) {
        throw "ExecuteExcel4Macro COM execution mismatch: expected 42, got $actual"
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    & cargo build -q -p oxidocs-cli --manifest-path (Join-Path $repoRoot 'Cargo.toml')
    if ($LASTEXITCODE -ne 0) {
        throw "oxidocs-cli build failed with exit code $LASTEXITCODE"
    }

    $analysis = (& $cliPath vba-analyze $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-analyze failed with exit code $LASTEXITCODE"
    }
    foreach ($expected in @(
        '[Excel4MacroProbe]',
        'verdict: B (data transformation): executes a legacy Excel 4.0 macro string',
        'procedures: 1, statements: 1, max nesting: 0, unparsed: 0',
        'formula engine: required',
        'Application.ExecuteExcel4Macro: executes a legacy Excel 4.0 macro string'
    )) {
        if (-not $analysis.Contains($expected)) {
            throw "Expected analysis fragment not found: $expected`n$analysis"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'Excel4MacroProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        -not $probeModule.needs_formula_engine -or
        $probeModule.api_names.'Application.ExecuteExcel4Macro' -ne 1) {
        throw 'ExecuteExcel4Macro JSON analysis did not preserve the XLM dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ExecuteExcel4Macro JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM ExecuteExcel4Macro returned: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA ExecuteExcel4Macro COM execution and analysis: PASS'
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    if (Test-Path -LiteralPath $probeRoot) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}
