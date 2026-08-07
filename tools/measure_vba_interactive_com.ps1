# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-interactive-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'interactive-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalInteractive = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalInteractive = [bool]$excel.Interactive
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'InteractiveProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableUserInput()
    Application.Interactive = False
End Sub

Public Sub EnableUserInput()
    Application.Interactive = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!InteractiveProbe.DisableUserInput")
    $disabled = [bool]$excel.Interactive
    if ($disabled) {
        throw 'Interactive COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!InteractiveProbe.EnableUserInput")
    $enabled = [bool]$excel.Interactive
    if (-not $enabled) {
        throw 'Interactive COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.Interactive = $originalInteractive
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
        '[InteractiveProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        "Application.Interactive: changes Excel's process-global keyboard and mouse input state"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'InteractiveProbe' })[0]
    $inputFindings = @($probeModule.findings | Where-Object { $_.reason -like '*keyboard and mouse input state*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.Interactive' -ne 2 -or
        $inputFindings.Count -ne 2) {
        throw 'Interactive JSON analysis did not preserve global input-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Interactive JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Interactive states: original=$originalInteractive disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA Interactive COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalInteractive) {
        try { $excel.Interactive = $originalInteractive } catch {}
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
