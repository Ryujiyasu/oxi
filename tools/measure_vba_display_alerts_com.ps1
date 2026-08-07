# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-alerts-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-alerts-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $true
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayAlertsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableAlerts()
    Application.DisplayAlerts = False
End Sub

Public Sub EnableAlerts()
    Application.DisplayAlerts = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayAlertsProbe.DisableAlerts")
    $disabled = [bool]$excel.DisplayAlerts
    if ($disabled) {
        throw 'DisplayAlerts COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayAlertsProbe.EnableAlerts")
    $enabled = [bool]$excel.DisplayAlerts
    if (-not $enabled) {
        throw 'DisplayAlerts COM enabled-state mismatch: expected True'
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
        '[DisplayAlertsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        "Application.DisplayAlerts: changes Excel's process-global alert handling and automatic default responses"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayAlertsProbe' })[0]
    $alertFindings = @($probeModule.findings | Where-Object { $_.reason -like '*automatic default responses*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayAlerts' -ne 2 -or
        $alertFindings.Count -ne 2) {
        throw 'DisplayAlerts JSON analysis did not preserve global alert-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'DisplayAlerts JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayAlerts states: disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA DisplayAlerts COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        try { $excel.DisplayAlerts = $true } catch {}
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
