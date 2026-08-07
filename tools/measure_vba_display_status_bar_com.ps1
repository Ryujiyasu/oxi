# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-status-bar-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-status-bar-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayStatusBar = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalDisplayStatusBar = [bool]$excel.DisplayStatusBar
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayStatusBarProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideStatusBar()
    Application.DisplayStatusBar = False
End Sub

Public Sub ShowStatusBar()
    Application.DisplayStatusBar = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayStatusBarProbe.HideStatusBar")
    $hidden = [bool]$excel.DisplayStatusBar
    if ($hidden) {
        throw 'DisplayStatusBar COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayStatusBarProbe.ShowStatusBar")
    $shown = [bool]$excel.DisplayStatusBar
    if (-not $shown) {
        throw 'DisplayStatusBar COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.DisplayStatusBar = $originalDisplayStatusBar
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
        '[DisplayStatusBarProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayStatusBar: shows or hides Excel's status-bar user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayStatusBarProbe' })[0]
    $displayFindings = @($probeModule.findings | Where-Object { $_.reason -like '*status-bar user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayStatusBar' -ne 2 -or
        $displayFindings.Count -ne 2) {
        throw 'DisplayStatusBar JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'DisplayStatusBar JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayStatusBar states: original=$originalDisplayStatusBar hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA DisplayStatusBar COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayStatusBar) {
        try { $excel.DisplayStatusBar = $originalDisplayStatusBar } catch {}
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
