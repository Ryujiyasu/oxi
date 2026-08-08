# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-workbook-tabs-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-workbook-tabs-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayWorkbookTabs = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayWorkbookTabs = [bool]$excel.ActiveWindow.DisplayWorkbookTabs

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowWorkbookTabsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideWorkbookTabs()
    Application.ActiveWindow.DisplayWorkbookTabs = False
End Sub

Public Sub ShowWorkbookTabs()
    Application.ActiveWindow.DisplayWorkbookTabs = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowWorkbookTabsProbe.HideWorkbookTabs")
    $hidden = [bool]$excel.ActiveWindow.DisplayWorkbookTabs
    if ($hidden) {
        throw 'ActiveWindow.DisplayWorkbookTabs COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowWorkbookTabsProbe.ShowWorkbookTabs")
    $shown = [bool]$excel.ActiveWindow.DisplayWorkbookTabs
    if (-not $shown) {
        throw 'ActiveWindow.DisplayWorkbookTabs COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.DisplayWorkbookTabs = $originalDisplayWorkbookTabs
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
        '[ActiveWindowWorkbookTabsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.ActiveWindow.DisplayWorkbookTabs: reads Excel's active UI context",
        'Application.ActiveWindow.DisplayWorkbookTabs: shows or hides workbook sheet tabs in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowWorkbookTabsProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $tabFindings = @($probeModule.findings | Where-Object { $_.reason -like '*workbook sheet tabs*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ActiveWindow.DisplayWorkbookTabs' -ne 2 -or
        $activeWindowFindings.Count -ne 2 -or
        $tabFindings.Count -ne 2) {
        throw 'ActiveWindow.DisplayWorkbookTabs JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.DisplayWorkbookTabs JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayWorkbookTabs states: original=$originalDisplayWorkbookTabs hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.DisplayWorkbookTabs COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayWorkbookTabs) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.DisplayWorkbookTabs = $originalDisplayWorkbookTabs
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
