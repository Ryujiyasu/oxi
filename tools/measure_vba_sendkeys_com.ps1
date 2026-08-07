# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-sendkeys-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'sendkeys-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.Caption = 'Oxi VBA AppActivate Probe'
    $workbook = $excel.Workbooks.Add()
    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'SendKeysProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ExerciseDesktopAutomation() As Long
    AppActivate Application.Caption
    SendKeys "{F15}", True
    ExerciseDesktopAutomation = 42
End Function
'@)

    $actual = [long]$excel.Run("'$($workbook.Name)'!SendKeysProbe.ExerciseDesktopAutomation")
    if ($actual -ne 42) {
        throw "AppActivate/SendKeys COM execution mismatch: expected 42, got $actual"
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
        '[SendKeysProbe]',
        'procedures: 1, statements: 3, max nesting: 0, unparsed: 0',
        '[C] AppActivate: activates a desktop application window by title',
        '[C] SendKeys: injects keystrokes into the active desktop application'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'SendKeysProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.class -ne 'C' -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.AppActivate -ne 1 -or
        $probeModule.api_names.SendKeys -ne 1) {
        throw 'AppActivate/SendKeys JSON analysis did not preserve desktop automation dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'AppActivate/SendKeys JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM desktop automation returned: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA AppActivate/SendKeys COM execution and analysis: PASS'
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
