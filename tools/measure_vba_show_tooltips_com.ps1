# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-show-tooltips-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'show-tooltips-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalShowToolTips = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalShowToolTips = [bool]$excel.ShowToolTips

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ShowToolTipsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideToolTips()
    Application.ShowToolTips = False
End Sub

Public Sub ShowToolTips()
    Application.ShowToolTips = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ShowToolTipsProbe.HideToolTips")
    $hidden = [bool]$excel.ShowToolTips
    if ($hidden) {
        throw 'Application.ShowToolTips COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!ShowToolTipsProbe.ShowToolTips")
    $shown = [bool]$excel.ShowToolTips
    if (-not $shown) {
        throw 'Application.ShowToolTips COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ShowToolTips = $originalShowToolTips
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
        '[ShowToolTipsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.ShowToolTips: shows or hides Excel's process-global tooltips user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ShowToolTipsProbe' })[0]
    $tooltipFindings = @($probeModule.findings | Where-Object { $_.reason -like '*tooltips user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ShowToolTips' -ne 2 -or
        $tooltipFindings.Count -ne 2) {
        throw 'Application.ShowToolTips JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.ShowToolTips JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM ShowToolTips states: original=$originalShowToolTips hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.ShowToolTips COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalShowToolTips) {
        try { $excel.ShowToolTips = $originalShowToolTips } catch {}
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
