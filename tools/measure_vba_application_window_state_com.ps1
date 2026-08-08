# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-application-window-state-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'application-window-state-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalWindowState = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalWindowState = [int]$excel.WindowState
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ApplicationWindowStateProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub MaximizeApplicationWindow()
    Application.WindowState = xlMaximized
End Sub

Public Sub NormalizeApplicationWindow()
    Application.WindowState = xlNormal
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ApplicationWindowStateProbe.MaximizeApplicationWindow")
    $maximized = [int]$excel.WindowState
    if ($maximized -ne -4137) {
        throw "Application.WindowState COM maximized-state mismatch: expected -4137, got $maximized"
    }

    $excel.Run("'$($workbook.Name)'!ApplicationWindowStateProbe.NormalizeApplicationWindow")
    $normal = [int]$excel.WindowState
    if ($normal -ne -4143) {
        throw "Application.WindowState COM normal-state mismatch: expected -4143, got $normal"
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.WindowState = $originalWindowState
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
        '[ApplicationWindowStateProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.WindowState: changes Excel's application-window minimized, normal, or maximized UI state"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ApplicationWindowStateProbe' })[0]
    $windowStateFindings = @($probeModule.findings | Where-Object { $_.reason -like '*application-window*UI state*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.WindowState' -ne 2 -or
        $windowStateFindings.Count -ne 2) {
        throw 'Application.WindowState JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.WindowState JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM WindowState values: original=$originalWindowState maximized=$maximized normal=$normal"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.WindowState COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalWindowState) {
        try { $excel.WindowState = $originalWindowState } catch {}
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
