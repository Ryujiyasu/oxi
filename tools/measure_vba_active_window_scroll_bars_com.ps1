# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-scroll-bars-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-scroll-bars-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalHorizontal = $null
$originalVertical = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalHorizontal = [bool]$excel.ActiveWindow.DisplayHorizontalScrollBar
    $originalVertical = [bool]$excel.ActiveWindow.DisplayVerticalScrollBar

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowScrollBarsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideScrollBars()
    Application.ActiveWindow.DisplayHorizontalScrollBar = False
    Application.ActiveWindow.DisplayVerticalScrollBar = False
End Sub

Public Sub ShowScrollBars()
    Application.ActiveWindow.DisplayHorizontalScrollBar = True
    Application.ActiveWindow.DisplayVerticalScrollBar = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowScrollBarsProbe.HideScrollBars")
    $hiddenHorizontal = [bool]$excel.ActiveWindow.DisplayHorizontalScrollBar
    $hiddenVertical = [bool]$excel.ActiveWindow.DisplayVerticalScrollBar
    if ($hiddenHorizontal -or $hiddenVertical) {
        throw "ActiveWindow scroll-bar COM hidden-state mismatch: horizontal=$hiddenHorizontal vertical=$hiddenVertical"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowScrollBarsProbe.ShowScrollBars")
    $shownHorizontal = [bool]$excel.ActiveWindow.DisplayHorizontalScrollBar
    $shownVertical = [bool]$excel.ActiveWindow.DisplayVerticalScrollBar
    if (-not $shownHorizontal -or -not $shownVertical) {
        throw "ActiveWindow scroll-bar COM shown-state mismatch: horizontal=$shownHorizontal vertical=$shownVertical"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.DisplayHorizontalScrollBar = $originalHorizontal
    $excel.ActiveWindow.DisplayVerticalScrollBar = $originalVertical
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
        '[ActiveWindowScrollBarsProbe]',
        'procedures: 2, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.ActiveWindow.DisplayHorizontalScrollBar: reads Excel's active UI context",
        'Application.ActiveWindow.DisplayHorizontalScrollBar: shows or hides the horizontal scroll bar in an Excel window',
        'Application.ActiveWindow.DisplayVerticalScrollBar: shows or hides the vertical scroll bar in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowScrollBarsProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $scrollBarFindings = @($probeModule.findings | Where-Object { $_.reason -like '*scroll bar in an Excel window*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ActiveWindow.DisplayHorizontalScrollBar' -ne 2 -or
        $probeModule.api_names.'Application.ActiveWindow.DisplayVerticalScrollBar' -ne 2 -or
        $activeWindowFindings.Count -ne 4 -or
        $scrollBarFindings.Count -ne 4) {
        throw 'ActiveWindow scroll-bar JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow scroll-bar JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM scroll bars: original=$originalHorizontal/$originalVertical hidden=$hiddenHorizontal/$hiddenVertical shown=$shownHorizontal/$shownVertical"
    $analysis.TrimEnd()
    Write-Output 'VBA active-window scroll-bar COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalHorizontal -and $null -ne $originalVertical) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.DisplayHorizontalScrollBar = $originalHorizontal
                $excel.ActiveWindow.DisplayVerticalScrollBar = $originalVertical
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
