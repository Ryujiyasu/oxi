# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-taskbar-windows-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'taskbar-windows-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalShowWindowsInTaskbar = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalShowWindowsInTaskbar = [bool]$excel.ShowWindowsInTaskbar
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'TaskbarWindowsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideTaskbarWindows()
    Application.ShowWindowsInTaskbar = False
End Sub

Public Sub ShowTaskbarWindows()
    Application.ShowWindowsInTaskbar = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!TaskbarWindowsProbe.HideTaskbarWindows")
    $afterHideRequest = [bool]$excel.ShowWindowsInTaskbar
    if (-not $afterHideRequest) {
        throw 'ShowWindowsInTaskbar COM behavior changed: current Excel was expected to ignore False'
    }

    $excel.Run("'$($workbook.Name)'!TaskbarWindowsProbe.ShowTaskbarWindows")
    $afterShowRequest = [bool]$excel.ShowWindowsInTaskbar
    if (-not $afterShowRequest) {
        throw 'ShowWindowsInTaskbar COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.ShowWindowsInTaskbar = $originalShowWindowsInTaskbar
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
        '[TaskbarWindowsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        'Application.ShowWindowsInTaskbar: requests showing or hiding Excel workbook windows in the Windows taskbar; modern Excel may ignore it'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'TaskbarWindowsProbe' })[0]
    $taskbarFindings = @($probeModule.findings | Where-Object { $_.reason -like '*Windows taskbar*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ShowWindowsInTaskbar' -ne 2 -or
        $taskbarFindings.Count -ne 2) {
        throw 'ShowWindowsInTaskbar JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ShowWindowsInTaskbar JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM ShowWindowsInTaskbar states: original=$originalShowWindowsInTaskbar afterFalse=$afterHideRequest afterTrue=$afterShowRequest"
    $analysis.TrimEnd()
    Write-Output 'VBA ShowWindowsInTaskbar COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalShowWindowsInTaskbar) {
        try { $excel.ShowWindowsInTaskbar = $originalShowWindowsInTaskbar } catch {}
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
