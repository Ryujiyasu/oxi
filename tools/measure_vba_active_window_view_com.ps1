# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-view-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-view-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalView = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalView = [int]$excel.ActiveWindow.View

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowViewProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub ShowPageBreakPreview()
    ActiveWindow.View = xlPageBreakPreview
End Sub

Public Sub ShowNormalView()
    Application.ActiveWindow.View = xlNormalView
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowViewProbe.ShowPageBreakPreview")
    $pageBreakPreview = [int]$excel.ActiveWindow.View
    if ($pageBreakPreview -ne 2) {
        throw "ActiveWindow.View COM page-break-preview mismatch: expected 2, got $pageBreakPreview"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowViewProbe.ShowNormalView")
    $normal = [int]$excel.ActiveWindow.View
    if ($normal -ne 1) {
        throw "ActiveWindow.View COM normal-view mismatch: expected 1, got $normal"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.View = $originalView
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
        '[ActiveWindowViewProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.View: reads Excel's active UI context",
        "ActiveWindow.View: changes Excel's active-window worksheet view mode",
        "Application.ActiveWindow.View: changes Excel's active-window worksheet view mode"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowViewProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $viewFindings = @($probeModule.findings | Where-Object { $_.reason -like '*worksheet view mode*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.View' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.View' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $viewFindings.Count -ne 2) {
        throw 'ActiveWindow.View JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.View JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM View values: original=$originalView pageBreakPreview=$pageBreakPreview normal=$normal"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.View COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalView) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.View = $originalView
            }
        }
        catch {}
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
