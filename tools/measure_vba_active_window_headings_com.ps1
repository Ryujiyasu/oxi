# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-headings-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-headings-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayHeadings = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayHeadings = [bool]$excel.ActiveWindow.DisplayHeadings

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowHeadingsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideHeadings()
    Application.ActiveWindow.DisplayHeadings = False
End Sub

Public Sub ShowHeadings()
    Application.ActiveWindow.DisplayHeadings = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowHeadingsProbe.HideHeadings")
    $hidden = [bool]$excel.ActiveWindow.DisplayHeadings
    if ($hidden) {
        throw 'ActiveWindow.DisplayHeadings COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowHeadingsProbe.ShowHeadings")
    $shown = [bool]$excel.ActiveWindow.DisplayHeadings
    if (-not $shown) {
        throw 'ActiveWindow.DisplayHeadings COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.DisplayHeadings = $originalDisplayHeadings
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
        '[ActiveWindowHeadingsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.ActiveWindow.DisplayHeadings: reads Excel's active UI context",
        'Application.ActiveWindow.DisplayHeadings: shows or hides worksheet row and column headings in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowHeadingsProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $headingFindings = @($probeModule.findings | Where-Object { $_.reason -like '*row and column headings*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ActiveWindow.DisplayHeadings' -ne 2 -or
        $activeWindowFindings.Count -ne 2 -or
        $headingFindings.Count -ne 2) {
        throw 'ActiveWindow.DisplayHeadings JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.DisplayHeadings JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayHeadings states: original=$originalDisplayHeadings hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.DisplayHeadings COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayHeadings) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.DisplayHeadings = $originalDisplayHeadings
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
