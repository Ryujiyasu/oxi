# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-enable-animations-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'enable-animations-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalEnableAnimations = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalEnableAnimations = [bool]$excel.EnableAnimations

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'EnableAnimationsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableAnimations()
    Application.EnableAnimations = False
End Sub

Public Sub EnableAnimations()
    Application.EnableAnimations = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!EnableAnimationsProbe.DisableAnimations")
    $disabled = [bool]$excel.EnableAnimations
    if ($disabled) {
        throw 'Application.EnableAnimations COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!EnableAnimationsProbe.EnableAnimations")
    $enabled = [bool]$excel.EnableAnimations
    if (-not $enabled) {
        throw 'Application.EnableAnimations COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.EnableAnimations = $originalEnableAnimations
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
        '[EnableAnimationsProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.EnableAnimations: changes Excel's process-global user-interface animation state"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'EnableAnimationsProbe' })[0]
    $animationFindings = @($probeModule.findings | Where-Object { $_.reason -like '*user-interface animation state*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.EnableAnimations' -ne 2 -or
        $animationFindings.Count -ne 2) {
        throw 'Application.EnableAnimations JSON analysis did not preserve its UI dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.EnableAnimations JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM EnableAnimations states: original=$originalEnableAnimations disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.EnableAnimations COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalEnableAnimations) {
        try { $excel.EnableAnimations = $originalEnableAnimations } catch {}
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
