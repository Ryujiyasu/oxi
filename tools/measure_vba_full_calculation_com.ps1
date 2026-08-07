# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-full-calculation-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'full-calculation-probe.xlsm'
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
    $component.Name = 'FullCalculationProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub RunCalculateFull()
    Application.CalculateFull
End Sub

Public Sub RunCalculateFullRebuild()
    Application.CalculateFullRebuild
End Sub
'@)

    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Range('A1').Value2 = 40
    $sheet.Range('A2').Formula = '=A1+2'
    $excel.Run("'$($workbook.Name)'!FullCalculationProbe.RunCalculateFull")
    $afterFull = [long]$sheet.Range('A2').Value2

    $sheet.Range('A1').Value2 = 41
    $excel.Run("'$($workbook.Name)'!FullCalculationProbe.RunCalculateFullRebuild")
    $afterRebuild = [long]$sheet.Range('A2').Value2
    if ($afterFull -ne 42 -or $afterRebuild -ne 43) {
        throw "Full calculation COM execution mismatch: full=$afterFull rebuild=$afterRebuild"
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
        '[FullCalculationProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'formula engine: required',
        'Application.CalculateFull: forces a full Excel workbook recalculation',
        "Application.CalculateFullRebuild: rebuilds Excel's formula dependencies and recalculates every workbook"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'FullCalculationProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        -not $probeModule.needs_formula_engine -or
        $probeModule.api_names.'Application.CalculateFull' -ne 1 -or
        $probeModule.api_names.'Application.CalculateFullRebuild' -ne 1) {
        throw 'Full calculation JSON analysis did not preserve recalculation dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Full calculation JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM recalculation values: full=$afterFull, rebuild=$afterRebuild"
    $analysis.TrimEnd()
    Write-Output 'VBA full calculation COM execution and analysis: PASS'
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
