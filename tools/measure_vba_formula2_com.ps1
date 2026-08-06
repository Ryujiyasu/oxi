# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-formula2-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'formula2-probe.xlsm'
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
    $component.Name = 'Formula2Probe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub WriteFormula2()
    With ThisWorkbook.Worksheets(1)
        .Range("A1").Value2 = 3
        .Range("B1").Formula2R1C1 = "=RC[-1]*2"
        .Range("D1").Formula2 = "=SUM(1,2)"
    End With
End Sub

Public Function Formula2Total() As Long
    Formula2Total = ThisWorkbook.Worksheets(1).Range("B1").Value2 + ThisWorkbook.Worksheets(1).Range("D1").Value2
End Function
'@)

    $excel.Run("'$($workbook.Name)'!Formula2Probe.WriteFormula2")
    $excel.Calculate()
    $actual = [long]$excel.Run("'$($workbook.Name)'!Formula2Probe.Formula2Total")
    if ($actual -ne 9) {
        throw "Formula2 COM execution mismatch: expected 9, got $actual"
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
        '[Formula2Probe]',
        'verdict: B (data transformation)',
        'procedures: 2, statements: 5, max nesting: 1, unparsed: 0',
        'formula engine: required',
        'Formula2: reads or writes dynamic-array formulas',
        'Formula2R1C1: reads or writes dynamic-array formulas in R1C1 notation'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'Formula2Probe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        -not $probeModule.needs_formula_engine -or
        $probeModule.api_names.'ThisWorkbook.Worksheets.Range.Formula2' -ne 1 -or
        $probeModule.api_names.'ThisWorkbook.Worksheets.Range.Formula2R1C1' -ne 1) {
        throw 'Formula2 JSON analysis did not preserve formula dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Formula2 JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM calculated Formula2 total: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Formula2 COM execution and analysis: PASS'
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
