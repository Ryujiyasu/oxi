# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-recent-files-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-recent-files-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayRecentFiles = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayRecentFiles = [bool]$excel.DisplayRecentFiles

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayRecentFilesProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideRecentFiles()
    Application.DisplayRecentFiles = False
End Sub

Public Sub ShowRecentFiles()
    Application.DisplayRecentFiles = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayRecentFilesProbe.HideRecentFiles")
    $hidden = [bool]$excel.DisplayRecentFiles
    if ($hidden) {
        throw 'Application.DisplayRecentFiles COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayRecentFilesProbe.ShowRecentFiles")
    $shown = [bool]$excel.DisplayRecentFiles
    if (-not $shown) {
        throw 'Application.DisplayRecentFiles COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.DisplayRecentFiles = $originalDisplayRecentFiles
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
        '[DisplayRecentFilesProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayRecentFiles: shows or hides Excel's process-global recent-files user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayRecentFilesProbe' })[0]
    $recentFilesFindings = @($probeModule.findings | Where-Object { $_.reason -like '*recent-files user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayRecentFiles' -ne 2 -or
        $recentFilesFindings.Count -ne 2) {
        throw 'Application.DisplayRecentFiles JSON analysis did not preserve its UI dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.DisplayRecentFiles JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayRecentFiles states: original=$originalDisplayRecentFiles hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.DisplayRecentFiles COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayRecentFiles) {
        try { $excel.DisplayRecentFiles = $originalDisplayRecentFiles } catch {}
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
