# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-ontime-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'ontime-probe.xlsm'
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
    $component.Name = 'OnTimeProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub ScheduleProbe()
    Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 1), Procedure:="'ontime-probe.xlsm'!OnTimeProbe.MarkComplete"
End Sub

Public Sub MarkComplete()
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = 42
End Sub
'@)

    $workbook.SaveAs($workbookPath, 52)
    $excel.Run("'$($workbook.Name)'!OnTimeProbe.ScheduleProbe")

    $actual = 0
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt 12) {
        Start-Sleep -Milliseconds 100
        try {
            $value = $workbook.Worksheets.Item(1).Range('A1').Value2
            if ($null -ne $value) {
                $actual = [long]$value
            }
        }
        catch {
            continue
        }
        if ($actual -eq 42) {
            break
        }
    }
    if ($actual -ne 42) {
        throw "Application.OnTime COM execution mismatch: expected 42, got $actual"
    }

    $workbook.Save()
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
    foreach ($expected in @(
        '[OnTimeProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        "Application.OnTime: schedules a macro by name; execution depends on Excel's application event loop"
    )) {
        if (-not $analysis.Contains($expected)) {
            throw "Expected analysis fragment not found: $expected`n$analysis"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'OnTimeProbe' })[0]
    $scheduledFindings = @($probeModule.findings | Where-Object { $_.reason -like '*application event loop*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.OnTime' -ne 1 -or
        $scheduledFindings.Count -ne 1) {
        throw 'Application.OnTime JSON analysis did not preserve scheduled dispatch diagnostics'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.OnTime JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Application.OnTime wrote: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.OnTime COM execution and analysis: PASS'
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
