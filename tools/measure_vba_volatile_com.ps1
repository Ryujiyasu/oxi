# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-volatile-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'volatile-probe.xlsm'
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
    $excel.Calculation = -4135
    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'VolatileProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Private invocationCount As Long

Public Function VolatileCounter() As Long
    Application.Volatile
    invocationCount = invocationCount + 1
    VolatileCounter = invocationCount
End Function
'@)

    $cell = $workbook.Worksheets.Item(1).Range('A1')
    $cell.Formula = '=VolatileCounter()'
    $excel.Calculate()
    $first = [long]$cell.Value2
    $excel.Calculate()
    $second = [long]$cell.Value2
    if ($first -lt 1 -or $second -ne ($first + 1)) {
        throw "Application.Volatile COM execution mismatch: first=$first second=$second"
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
    foreach ($expectedFragment in @(
        '[VolatileProbe]',
        'verdict: B (data transformation): marks a VBA function for execution on every Excel recalculation',
        'procedures: 1, statements: 3, max nesting: 0, unparsed: 0',
        'formula engine: required',
        'Application.Volatile: marks a VBA function for execution on every Excel recalculation'
    )) {
        if (-not $analysis.Contains($expectedFragment)) {
            throw "Expected analysis fragment not found: $expectedFragment`n$analysis"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'VolatileProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        -not $probeModule.needs_formula_engine -or
        $probeModule.api_names.'Application.Volatile' -ne 1) {
        throw 'Application.Volatile JSON analysis did not preserve recalculation dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.Volatile JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM volatile UDF counts: first=$first, second=$second"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.Volatile COM execution and analysis: PASS'
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
