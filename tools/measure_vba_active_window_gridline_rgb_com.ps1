# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-gridline-rgb-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-gridline-rgb-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalColor = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalColor = [int]$excel.ActiveWindow.GridlineColor

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowGridlineRgbProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub SetRedGridlinesByRgb()
    ActiveWindow.GridlineColor = RGB(255, 0, 0)
End Sub

Public Sub SetBlueGridlinesByRgb()
    Application.ActiveWindow.GridlineColor = RGB(0, 0, 255)
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowGridlineRgbProbe.SetRedGridlinesByRgb")
    $redColor = [int]$excel.ActiveWindow.GridlineColor
    $redIndex = [int]$excel.ActiveWindow.GridlineColorIndex
    if ($redColor -ne 255 -or $redIndex -ne 3) {
        throw "ActiveWindow.GridlineColor COM red mismatch: color=$redColor index=$redIndex"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowGridlineRgbProbe.SetBlueGridlinesByRgb")
    $blueColor = [int]$excel.ActiveWindow.GridlineColor
    $blueIndex = [int]$excel.ActiveWindow.GridlineColorIndex
    if ($blueColor -ne 16711680 -or $blueIndex -ne 5) {
        throw "ActiveWindow.GridlineColor COM blue mismatch: color=$blueColor index=$blueIndex"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.GridlineColor = $originalColor
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
        '[ActiveWindowGridlineRgbProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.GridlineColor: reads Excel's active UI context",
        'ActiveWindow.GridlineColor: changes the worksheet gridline colour using the workbook colour palette in an Excel window',
        'Application.ActiveWindow.GridlineColor: changes the worksheet gridline colour using the workbook colour palette in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowGridlineRgbProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $gridlineColorFindings = @($probeModule.findings | Where-Object { $_.reason -like '*workbook colour palette*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.GridlineColor' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.GridlineColor' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $gridlineColorFindings.Count -ne 2) {
        throw 'ActiveWindow.GridlineColor JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.GridlineColor JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM gridline RGB values: original=$originalColor red=$redColor redIndex=$redIndex blue=$blueColor blueIndex=$blueIndex"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.GridlineColor COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalColor) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.GridlineColor = $originalColor
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
