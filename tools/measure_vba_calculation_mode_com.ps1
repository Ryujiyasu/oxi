# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-calculation-mode-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'calculation-mode-probe.xlsm'
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
    $component.Name = 'CalculationModeProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub SetManualCalculation()
    Application.Calculation = xlCalculationManual
End Sub

Public Sub SetAutomaticCalculation()
    Application.Calculation = xlCalculationAutomatic
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!CalculationModeProbe.SetManualCalculation")
    $manual = [int]$excel.Calculation
    if ($manual -ne -4135) {
        throw "Calculation COM manual mode mismatch: expected -4135, got $manual"
    }

    $excel.Run("'$($workbook.Name)'!CalculationModeProbe.SetAutomaticCalculation")
    $automatic = [int]$excel.Calculation
    if ($automatic -ne -4105) {
        throw "Calculation COM automatic mode mismatch: expected -4105, got $automatic"
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
        '[CalculationModeProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'formula engine: required',
        "Application.Calculation: changes Excel's process-global automatic calculation mode"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'CalculationModeProbe' })[0]
    $modeFindings = @($probeModule.findings | Where-Object { $_.reason -like '*automatic calculation mode*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        -not $probeModule.needs_formula_engine -or
        $probeModule.api_names.'Application.Calculation' -ne 2 -or
        $modeFindings.Count -ne 2) {
        throw 'Calculation JSON analysis did not preserve formula-engine and global-mode dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Calculation JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM calculation constants: manual=$manual automatic=$automatic"
    $analysis.TrimEnd()
    Write-Output 'VBA Calculation COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        try { $excel.Calculation = -4105 } catch {}
    }
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
