# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-application-wait-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'application-wait-probe.xlsm'
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
    $component.Name = 'ApplicationWaitProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function WaitOneSecond() As Boolean
    WaitOneSecond = Application.Wait(Now + TimeSerial(0, 0, 1))
End Function
'@)

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $actual = [bool]$excel.Run("'$($workbook.Name)'!ApplicationWaitProbe.WaitOneSecond")
    $stopwatch.Stop()
    $elapsedSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 3)
    if (-not $actual) {
        throw 'Application.Wait COM return mismatch: expected True'
    }
    if ($elapsedSeconds -lt 0.5 -or $elapsedSeconds -gt 5.0) {
        throw "Application.Wait COM duration mismatch: expected about one second, got $elapsedSeconds seconds"
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
        '[ApplicationWaitProbe]',
        'procedures: 1, statements: 1, max nesting: 0, unparsed: 0',
        'Application.Wait: blocks Excel until a wall-clock deadline and suspends application activity'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ApplicationWaitProbe' })[0]
    $waitFindings = @($probeModule.findings | Where-Object { $_.reason -like '*wall-clock deadline*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.Wait' -ne 1 -or
        $waitFindings.Count -ne 1) {
        throw 'Application.Wait JSON analysis did not preserve blocking clock dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.Wait JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Application.Wait: returned=$actual elapsed=${elapsedSeconds}s"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.Wait COM execution and analysis: PASS'
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
