# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-outline-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-outline-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayOutline = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayOutline = [bool]$excel.ActiveWindow.DisplayOutline

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowOutlineProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideOutlineSymbols()
    ActiveWindow.DisplayOutline = False
End Sub

Public Sub ShowOutlineSymbols()
    Application.ActiveWindow.DisplayOutline = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowOutlineProbe.HideOutlineSymbols")
    $hidden = [bool]$excel.ActiveWindow.DisplayOutline
    if ($hidden) {
        throw 'ActiveWindow.DisplayOutline COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowOutlineProbe.ShowOutlineSymbols")
    $shown = [bool]$excel.ActiveWindow.DisplayOutline
    if (-not $shown) {
        throw 'ActiveWindow.DisplayOutline COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.DisplayOutline = $originalDisplayOutline
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
        '[ActiveWindowOutlineProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.DisplayOutline: reads Excel's active UI context",
        'ActiveWindow.DisplayOutline: shows or hides worksheet outline symbols in an Excel window',
        'Application.ActiveWindow.DisplayOutline: shows or hides worksheet outline symbols in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowOutlineProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $outlineFindings = @($probeModule.findings | Where-Object { $_.reason -like '*worksheet outline symbols*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.DisplayOutline' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.DisplayOutline' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $outlineFindings.Count -ne 2) {
        throw 'ActiveWindow.DisplayOutline JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.DisplayOutline JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayOutline states: original=$originalDisplayOutline hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.DisplayOutline COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayOutline) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.DisplayOutline = $originalDisplayOutline
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
