# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-tab-ratio-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-tab-ratio-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalTabRatio = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalTabRatio = [double]$excel.ActiveWindow.TabRatio

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowTabRatioProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub AllocateQuarterToTabs()
    ActiveWindow.TabRatio = 0.25
End Sub

Public Sub AllocateThreeQuartersToTabs()
    Application.ActiveWindow.TabRatio = 0.75
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowTabRatioProbe.AllocateQuarterToTabs")
    $quarter = [double]$excel.ActiveWindow.TabRatio
    if ([Math]::Abs($quarter - 0.25) -gt 0.000001) {
        throw "ActiveWindow.TabRatio COM quarter mismatch: expected 0.25, got $quarter"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowTabRatioProbe.AllocateThreeQuartersToTabs")
    $threeQuarters = [double]$excel.ActiveWindow.TabRatio
    if ([Math]::Abs($threeQuarters - 0.75) -gt 0.000001) {
        throw "ActiveWindow.TabRatio COM three-quarter mismatch: expected 0.75, got $threeQuarters"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.TabRatio = $originalTabRatio
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
        '[ActiveWindowTabRatioProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.TabRatio: reads Excel's active UI context",
        'ActiveWindow.TabRatio: changes the width allocation between workbook tabs and the horizontal scroll bar',
        'Application.ActiveWindow.TabRatio: changes the width allocation between workbook tabs and the horizontal scroll bar'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowTabRatioProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $tabRatioFindings = @($probeModule.findings | Where-Object { $_.reason -like '*width allocation*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.TabRatio' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.TabRatio' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $tabRatioFindings.Count -ne 2) {
        throw 'ActiveWindow.TabRatio JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.TabRatio JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM TabRatio values: original=$originalTabRatio quarter=$quarter threeQuarters=$threeQuarters"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.TabRatio COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalTabRatio) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.TabRatio = $originalTabRatio
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
