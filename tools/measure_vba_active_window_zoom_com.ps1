# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-zoom-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-zoom-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalZoom = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalZoom = [int]$excel.ActiveWindow.Zoom

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowZoomProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub SetZoom75()
    ActiveWindow.Zoom = 75
End Sub

Public Sub SetZoom100()
    Application.ActiveWindow.Zoom = 100
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowZoomProbe.SetZoom75")
    $zoom75 = [int]$excel.ActiveWindow.Zoom
    if ($zoom75 -ne 75) {
        throw "ActiveWindow.Zoom COM 75-percent mismatch: expected 75, got $zoom75"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowZoomProbe.SetZoom100")
    $zoom100 = [int]$excel.ActiveWindow.Zoom
    if ($zoom100 -ne 100) {
        throw "ActiveWindow.Zoom COM 100-percent mismatch: expected 100, got $zoom100"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.Zoom = $originalZoom
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
        '[ActiveWindowZoomProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.Zoom: reads Excel's active UI context",
        "ActiveWindow.Zoom: changes the zoom level of Excel's active window",
        "Application.ActiveWindow.Zoom: changes the zoom level of Excel's active window"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowZoomProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $zoomFindings = @($probeModule.findings | Where-Object { $_.reason -like '*zoom level*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.Zoom' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.Zoom' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $zoomFindings.Count -ne 2) {
        throw 'ActiveWindow.Zoom JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.Zoom JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Zoom values: original=$originalZoom zoom75=$zoom75 zoom100=$zoom100"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.Zoom COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalZoom) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.Zoom = $originalZoom
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
