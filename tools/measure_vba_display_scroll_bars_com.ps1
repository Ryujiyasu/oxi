# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-scroll-bars-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-scroll-bars-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayScrollBars = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayScrollBars = [bool]$excel.DisplayScrollBars

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayScrollBarsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideScrollBars()
    Application.DisplayScrollBars = False
End Sub

Public Sub ShowScrollBars()
    Application.DisplayScrollBars = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayScrollBarsProbe.HideScrollBars")
    $hidden = [bool]$excel.DisplayScrollBars
    if ($hidden) {
        throw 'Application.DisplayScrollBars COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayScrollBarsProbe.ShowScrollBars")
    $shown = [bool]$excel.DisplayScrollBars
    if (-not $shown) {
        throw 'Application.DisplayScrollBars COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.DisplayScrollBars = $originalDisplayScrollBars
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
        '[DisplayScrollBarsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayScrollBars: shows or hides Excel's process-global scroll-bar user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayScrollBarsProbe' })[0]
    $scrollBarFindings = @($probeModule.findings | Where-Object { $_.reason -like '*scroll-bar user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayScrollBars' -ne 2 -or
        $scrollBarFindings.Count -ne 2) {
        throw 'Application.DisplayScrollBars JSON analysis did not preserve its UI dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.DisplayScrollBars JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayScrollBars states: original=$originalDisplayScrollBars hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.DisplayScrollBars COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayScrollBars) {
        try { $excel.DisplayScrollBars = $originalDisplayScrollBars } catch {}
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
