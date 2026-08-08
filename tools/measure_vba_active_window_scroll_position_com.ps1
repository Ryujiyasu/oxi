# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-scroll-position-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-scroll-position-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalScrollRow = $null
$originalScrollColumn = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalScrollRow = [int]$excel.ActiveWindow.ScrollRow
    $originalScrollColumn = [int]$excel.ActiveWindow.ScrollColumn

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowScrollPositionProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub ScrollAwayFromOrigin()
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollColumn = 5
End Sub

Public Sub ScrollToOrigin()
    Application.ActiveWindow.ScrollRow = 1
    Application.ActiveWindow.ScrollColumn = 1
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowScrollPositionProbe.ScrollAwayFromOrigin")
    $scrolledRow = [int]$excel.ActiveWindow.ScrollRow
    $scrolledColumn = [int]$excel.ActiveWindow.ScrollColumn
    if ($scrolledRow -ne 10 -or $scrolledColumn -ne 5) {
        throw "ActiveWindow scroll-position COM mismatch: row=$scrolledRow column=$scrolledColumn"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowScrollPositionProbe.ScrollToOrigin")
    $originRow = [int]$excel.ActiveWindow.ScrollRow
    $originColumn = [int]$excel.ActiveWindow.ScrollColumn
    if ($originRow -ne 1 -or $originColumn -ne 1) {
        throw "ActiveWindow origin-position COM mismatch: row=$originRow column=$originColumn"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.ScrollRow = $originalScrollRow
    $excel.ActiveWindow.ScrollColumn = $originalScrollColumn
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
        '[ActiveWindowScrollPositionProbe]',
        'procedures: 2, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.ScrollRow: reads Excel's active UI context",
        'ActiveWindow.ScrollRow: changes the first visible row in an Excel window',
        'Application.ActiveWindow.ScrollColumn: changes the first visible column in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowScrollPositionProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $scrollFindings = @($probeModule.findings | Where-Object { $_.reason -like '*first visible*' })
    $expectedNames = @(
        'ActiveWindow.ScrollRow',
        'ActiveWindow.ScrollColumn',
        'Application.ActiveWindow.ScrollRow',
        'Application.ActiveWindow.ScrollColumn'
    )
    foreach ($expectedName in $expectedNames) {
        if ($probeModule.api_names.$expectedName -ne 1) {
            throw "ActiveWindow scroll-position JSON API count mismatch: $expectedName"
        }
    }
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $activeWindowFindings.Count -ne 4 -or
        $scrollFindings.Count -ne 4) {
        throw 'ActiveWindow scroll-position JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow scroll-position JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM scroll position: original=$originalScrollRow/$originalScrollColumn scrolled=$scrolledRow/$scrolledColumn origin=$originRow/$originColumn"
    $analysis.TrimEnd()
    Write-Output 'VBA active-window scroll-position COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalScrollRow -and $null -ne $originalScrollColumn) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.ScrollRow = $originalScrollRow
                $excel.ActiveWindow.ScrollColumn = $originalScrollColumn
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
