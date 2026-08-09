# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-gridline-color-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-gridline-color-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalColorIndex = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalColorIndex = [int]$excel.ActiveWindow.GridlineColorIndex

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowGridlineColorProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub SetRedGridlines()
    ActiveWindow.GridlineColorIndex = 3
End Sub

Public Sub SetBlueGridlines()
    Application.ActiveWindow.GridlineColorIndex = 5
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowGridlineColorProbe.SetRedGridlines")
    $redIndex = [int]$excel.ActiveWindow.GridlineColorIndex
    $redColor = [int]$excel.ActiveWindow.GridlineColor
    if ($redIndex -ne 3 -or $redColor -ne 255) {
        throw "ActiveWindow.GridlineColorIndex COM red mismatch: index=$redIndex color=$redColor"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowGridlineColorProbe.SetBlueGridlines")
    $blueIndex = [int]$excel.ActiveWindow.GridlineColorIndex
    $blueColor = [int]$excel.ActiveWindow.GridlineColor
    if ($blueIndex -ne 5 -or $blueColor -ne 16711680) {
        throw "ActiveWindow.GridlineColorIndex COM blue mismatch: index=$blueIndex color=$blueColor"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.GridlineColorIndex = $originalColorIndex
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
        '[ActiveWindowGridlineColorProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.GridlineColorIndex: reads Excel's active UI context",
        'ActiveWindow.GridlineColorIndex: changes the worksheet gridline colour in an Excel window',
        'Application.ActiveWindow.GridlineColorIndex: changes the worksheet gridline colour in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowGridlineColorProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $gridlineColorFindings = @($probeModule.findings | Where-Object { $_.reason -like '*gridline colour*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.GridlineColorIndex' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.GridlineColorIndex' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $gridlineColorFindings.Count -ne 2) {
        throw 'ActiveWindow.GridlineColorIndex JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.GridlineColorIndex JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM gridline colors: originalIndex=$originalColorIndex redIndex=$redIndex redColor=$redColor blueIndex=$blueIndex blueColor=$blueColor"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.GridlineColorIndex COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalColorIndex) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.GridlineColorIndex = $originalColorIndex
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
