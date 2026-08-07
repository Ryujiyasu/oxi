# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-screen-updating-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'screen-updating-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalScreenUpdating = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalScreenUpdating = [bool]$excel.ScreenUpdating
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ScreenUpdatingProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableScreenUpdating()
    Application.ScreenUpdating = False
End Sub

Public Sub EnableScreenUpdating()
    Application.ScreenUpdating = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ScreenUpdatingProbe.DisableScreenUpdating")
    $disabled = [bool]$excel.ScreenUpdating
    if ($disabled) {
        throw 'ScreenUpdating COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!ScreenUpdatingProbe.EnableScreenUpdating")
    $enabled = [bool]$excel.ScreenUpdating
    if (-not $enabled) {
        throw 'ScreenUpdating COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.ScreenUpdating = $originalScreenUpdating
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
        '[ScreenUpdatingProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        "Application.ScreenUpdating: changes Excel's process-global screen redraw state"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ScreenUpdatingProbe' })[0]
    $redrawFindings = @($probeModule.findings | Where-Object { $_.reason -like '*screen redraw state*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ScreenUpdating' -ne 2 -or
        $redrawFindings.Count -ne 2) {
        throw 'ScreenUpdating JSON analysis did not preserve global redraw-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ScreenUpdating JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM ScreenUpdating states: original=$originalScreenUpdating disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA ScreenUpdating COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalScreenUpdating) {
        try { $excel.ScreenUpdating = $originalScreenUpdating } catch {}
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
