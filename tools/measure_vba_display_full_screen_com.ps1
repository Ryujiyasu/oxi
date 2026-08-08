# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-full-screen-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-full-screen-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayFullScreen = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalDisplayFullScreen = [bool]$excel.DisplayFullScreen
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayFullScreenProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub EnterFullScreen()
    Application.DisplayFullScreen = True
End Sub

Public Sub LeaveFullScreen()
    Application.DisplayFullScreen = False
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayFullScreenProbe.EnterFullScreen")
    $entered = [bool]$excel.DisplayFullScreen
    if (-not $entered) {
        throw 'DisplayFullScreen COM entered-state mismatch: expected True'
    }

    $excel.Run("'$($workbook.Name)'!DisplayFullScreenProbe.LeaveFullScreen")
    $left = [bool]$excel.DisplayFullScreen
    if ($left) {
        throw 'DisplayFullScreen COM left-state mismatch: expected False'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.DisplayFullScreen = $originalDisplayFullScreen
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
        '[DisplayFullScreenProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayFullScreen: switches Excel's application window into or out of full-screen mode"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayFullScreenProbe' })[0]
    $fullScreenFindings = @($probeModule.findings | Where-Object { $_.reason -like '*full-screen mode*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayFullScreen' -ne 2 -or
        $fullScreenFindings.Count -ne 2) {
        throw 'DisplayFullScreen JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'DisplayFullScreen JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayFullScreen states: original=$originalDisplayFullScreen entered=$entered left=$left"
    $analysis.TrimEnd()
    Write-Output 'VBA DisplayFullScreen COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayFullScreen) {
        try { $excel.DisplayFullScreen = $originalDisplayFullScreen } catch {}
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
