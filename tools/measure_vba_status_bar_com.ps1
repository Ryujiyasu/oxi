# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-status-bar-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'status-bar-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'StatusBarProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ShowProgress() As String
    Application.StatusBar = "Oxi 42"
    ShowProgress = CStr(Application.StatusBar)
    Application.StatusBar = False
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!StatusBarProbe.ShowProgress")
    if ($actual -ne 'Oxi 42') {
        throw "StatusBar COM display mismatch: expected 'Oxi 42', got '$actual'"
    }
    $releasedStatus = $excel.StatusBar
    if ($releasedStatus -ne $false) {
        $releasedType = if ($null -eq $releasedStatus) { '<null>' } else { $releasedStatus.GetType().FullName }
        throw "StatusBar COM release mismatch: expected False, got '$releasedStatus' ($releasedType)"
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
        '[StatusBarProbe]',
        'procedures: 1, statements: 3, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.StatusBar: writes progress or status text to Excel's user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'StatusBarProbe' })[0]
    $statusFindings = @($probeModule.findings | Where-Object { $_.reason -like '*status text*user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.StatusBar' -ne 3 -or
        $statusFindings.Count -ne 3) {
        throw 'StatusBar JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'StatusBar JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM StatusBar text: $actual; released=$($releasedStatus -eq $false)"
    $analysis.TrimEnd()
    Write-Output 'VBA StatusBar COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        try { $excel.StatusBar = $false } catch {}
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
