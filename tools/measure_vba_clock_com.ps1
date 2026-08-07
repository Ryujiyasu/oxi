# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-clock-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'clock-probe.xlsm'
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
    $component.Name = 'ClockProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function CurrentTimestamp() As Date
    CurrentTimestamp = Now
End Function

Public Function SecondsSinceMidnight() As Double
    SecondsSinceMidnight = Timer
End Function
'@)

    $before = Get-Date
    $actualTimestamp = [datetime]$excel.Run("'$($workbook.Name)'!ClockProbe.CurrentTimestamp")
    $actualTimer = [double]$excel.Run("'$($workbook.Name)'!ClockProbe.SecondsSinceMidnight")
    $after = Get-Date
    if ($actualTimestamp -lt $before.AddSeconds(-2) -or $actualTimestamp -gt $after.AddSeconds(2)) {
        throw "Now COM execution mismatch: before=$before actual=$actualTimestamp after=$after"
    }
    $expectedTimer = $after.TimeOfDay.TotalSeconds
    $timerDifference = [math]::Abs($actualTimer - $expectedTimer)
    $timerDifference = [math]::Min($timerDifference, 86400 - $timerDifference)
    if ($timerDifference -gt 3) {
        throw "Timer COM execution mismatch: expected about $expectedTimer, got $actualTimer"
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
        '[ClockProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'Now: reads the system clock; results depend on execution time and local time zone',
        'Timer: reads the system clock; results depend on execution time and local time zone'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ClockProbe' })[0]
    $clockFindings = @($probeModule.findings | Where-Object { $_.reason -like '*local time zone*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.Now -ne 1 -or
        $probeModule.api_names.Timer -ne 1 -or
        $clockFindings.Count -ne 2) {
        throw 'Clock JSON analysis did not preserve time-dependent behavior'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Clock JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM clock: Now=$actualTimestamp, Timer=$actualTimer"
    $analysis.TrimEnd()
    Write-Output 'VBA clock COM execution and analysis: PASS'
}
finally {
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
