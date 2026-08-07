# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-automation-security-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'automation-security-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalSecurity = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalSecurity = [int]$excel.AutomationSecurity
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'AutomationSecurityProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableAutomatedMacros()
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
End Sub

Public Sub AllowAutomatedMacros()
    Application.AutomationSecurity = msoAutomationSecurityLow
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!AutomationSecurityProbe.DisableAutomatedMacros")
    $forceDisabled = [int]$excel.AutomationSecurity
    if ($forceDisabled -ne 3) {
        throw "AutomationSecurity COM force-disable mismatch: expected 3, got $forceDisabled"
    }

    $excel.Run("'$($workbook.Name)'!AutomationSecurityProbe.AllowAutomatedMacros")
    $low = [int]$excel.AutomationSecurity
    if ($low -ne 1) {
        throw "AutomationSecurity COM low-security mismatch: expected 1, got $low"
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.AutomationSecurity = $originalSecurity
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
        '[AutomationSecurityProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        "Application.AutomationSecurity: changes Excel's process-global macro policy for programmatically opened files"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'AutomationSecurityProbe' })[0]
    $securityFindings = @($probeModule.findings | Where-Object { $_.reason -like '*programmatically opened files*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.AutomationSecurity' -ne 2 -or
        $securityFindings.Count -ne 2) {
        throw 'AutomationSecurity JSON analysis did not preserve global macro-policy dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'AutomationSecurity JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM AutomationSecurity values: original=$originalSecurity force-disabled=$forceDisabled low=$low"
    $analysis.TrimEnd()
    Write-Output 'VBA AutomationSecurity COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalSecurity) {
        try { $excel.AutomationSecurity = $originalSecurity } catch {}
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
