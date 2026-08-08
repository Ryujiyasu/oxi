# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-freeze-panes-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-freeze-panes-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalSplitRow = $null
$originalSplitColumn = $null
$originalFreezePanes = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalSplitRow = [int]$excel.ActiveWindow.SplitRow
    $originalSplitColumn = [int]$excel.ActiveWindow.SplitColumn
    $originalFreezePanes = [bool]$excel.ActiveWindow.FreezePanes

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowFreezePanesProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub FreezeFirstRowAndColumn()
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 1
    ActiveWindow.FreezePanes = True
End Sub

Public Sub ClearFrozenPanes()
    Application.ActiveWindow.FreezePanes = False
    Application.ActiveWindow.SplitRow = 0
    Application.ActiveWindow.SplitColumn = 0
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowFreezePanesProbe.FreezeFirstRowAndColumn")
    $frozen = [bool]$excel.ActiveWindow.FreezePanes
    $frozenRow = [int]$excel.ActiveWindow.SplitRow
    $frozenColumn = [int]$excel.ActiveWindow.SplitColumn
    if (-not $frozen -or $frozenRow -ne 1 -or $frozenColumn -ne 1) {
        throw "ActiveWindow freeze-pane COM state mismatch: frozen=$frozen row=$frozenRow column=$frozenColumn"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowFreezePanesProbe.ClearFrozenPanes")
    $cleared = [bool]$excel.ActiveWindow.FreezePanes
    $clearedRow = [int]$excel.ActiveWindow.SplitRow
    $clearedColumn = [int]$excel.ActiveWindow.SplitColumn
    if ($cleared -or $clearedRow -ne 0 -or $clearedColumn -ne 0) {
        throw "ActiveWindow cleared-pane COM state mismatch: frozen=$cleared row=$clearedRow column=$clearedColumn"
    }

    $workbook.SaveAs($workbookPath, 52)
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
        '[ActiveWindowFreezePanesProbe]',
        'procedures: 2, statements: 6, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.SplitRow: reads Excel's active UI context",
        'ActiveWindow.SplitRow: changes the pane layout of an Excel window',
        'Application.ActiveWindow.FreezePanes: changes the pane layout of an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowFreezePanesProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $paneFindings = @($probeModule.findings | Where-Object { $_.reason -like '*pane layout*' })
    $expectedNames = @(
        'ActiveWindow.SplitRow',
        'ActiveWindow.SplitColumn',
        'ActiveWindow.FreezePanes',
        'Application.ActiveWindow.FreezePanes',
        'Application.ActiveWindow.SplitRow',
        'Application.ActiveWindow.SplitColumn'
    )
    foreach ($expectedName in $expectedNames) {
        if ($probeModule.api_names.$expectedName -ne 1) {
            throw "ActiveWindow pane JSON API count mismatch: $expectedName"
        }
    }
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $activeWindowFindings.Count -ne 6 -or
        $paneFindings.Count -ne 6) {
        throw 'ActiveWindow pane JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow pane JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM panes: original=$originalFreezePanes/$originalSplitRow/$originalSplitColumn frozen=$frozen/$frozenRow/$frozenColumn cleared=$cleared/$clearedRow/$clearedColumn"
    $analysis.TrimEnd()
    Write-Output 'VBA active-window freeze-pane COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalFreezePanes) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.FreezePanes = $false
                $excel.ActiveWindow.SplitRow = $originalSplitRow
                $excel.ActiveWindow.SplitColumn = $originalSplitColumn
                $excel.ActiveWindow.FreezePanes = $originalFreezePanes
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
