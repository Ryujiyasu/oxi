# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-formula-bar-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-formula-bar-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayFormulaBar = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalDisplayFormulaBar = [bool]$excel.DisplayFormulaBar
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayFormulaBarProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideFormulaBar()
    Application.DisplayFormulaBar = False
End Sub

Public Sub ShowFormulaBar()
    Application.DisplayFormulaBar = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayFormulaBarProbe.HideFormulaBar")
    $hidden = [bool]$excel.DisplayFormulaBar
    if ($hidden) {
        throw 'DisplayFormulaBar COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayFormulaBarProbe.ShowFormulaBar")
    $shown = [bool]$excel.DisplayFormulaBar
    if (-not $shown) {
        throw 'DisplayFormulaBar COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.DisplayFormulaBar = $originalDisplayFormulaBar
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
        '[DisplayFormulaBarProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayFormulaBar: shows or hides Excel's formula-bar user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayFormulaBarProbe' })[0]
    $formulaBarFindings = @($probeModule.findings | Where-Object { $_.reason -like '*formula-bar user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayFormulaBar' -ne 2 -or
        $formulaBarFindings.Count -ne 2) {
        throw 'DisplayFormulaBar JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'DisplayFormulaBar JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayFormulaBar states: original=$originalDisplayFormulaBar hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA DisplayFormulaBar COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayFormulaBar) {
        try { $excel.DisplayFormulaBar = $originalDisplayFormulaBar } catch {}
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
