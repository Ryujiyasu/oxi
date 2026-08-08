# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-value-display-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-value-display-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayFormulas = $null
$originalDisplayZeros = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayFormulas = [bool]$excel.ActiveWindow.DisplayFormulas
    $originalDisplayZeros = [bool]$excel.ActiveWindow.DisplayZeros

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowValueDisplayProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub ShowFormulasAndHideZeros()
    ActiveWindow.DisplayFormulas = True
    ActiveWindow.DisplayZeros = False
End Sub

Public Sub ShowValuesAndZeros()
    Application.ActiveWindow.DisplayFormulas = False
    Application.ActiveWindow.DisplayZeros = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowValueDisplayProbe.ShowFormulasAndHideZeros")
    $formulaMode = [bool]$excel.ActiveWindow.DisplayFormulas
    $zerosHidden = [bool]$excel.ActiveWindow.DisplayZeros
    if (-not $formulaMode -or $zerosHidden) {
        throw "ActiveWindow value-display COM configured-state mismatch: formulas=$formulaMode zeros=$zerosHidden"
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowValueDisplayProbe.ShowValuesAndZeros")
    $valueMode = [bool]$excel.ActiveWindow.DisplayFormulas
    $zerosShown = [bool]$excel.ActiveWindow.DisplayZeros
    if ($valueMode -or -not $zerosShown) {
        throw "ActiveWindow value-display COM restored-state mismatch: formulas=$valueMode zeros=$zerosShown"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.DisplayFormulas = $originalDisplayFormulas
    $excel.ActiveWindow.DisplayZeros = $originalDisplayZeros
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
        '[ActiveWindowValueDisplayProbe]',
        'procedures: 2, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.DisplayFormulas: reads Excel's active UI context",
        'ActiveWindow.DisplayFormulas: switches an Excel window between formula and calculated-value display',
        'Application.ActiveWindow.DisplayZeros: shows or hides zero values in an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowValueDisplayProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $valueDisplayFindings = @($probeModule.findings | Where-Object {
        $_.reason -like '*calculated-value display*' -or $_.reason -like '*zero values*'
    })
    $expectedNames = @(
        'ActiveWindow.DisplayFormulas',
        'ActiveWindow.DisplayZeros',
        'Application.ActiveWindow.DisplayFormulas',
        'Application.ActiveWindow.DisplayZeros'
    )
    foreach ($expectedName in $expectedNames) {
        if ($probeModule.api_names.$expectedName -ne 1) {
            throw "ActiveWindow value-display JSON API count mismatch: $expectedName"
        }
    }
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $activeWindowFindings.Count -ne 4 -or
        $valueDisplayFindings.Count -ne 4) {
        throw 'ActiveWindow value-display JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow value-display JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM value display: original=$originalDisplayFormulas/$originalDisplayZeros configured=$formulaMode/$zerosHidden restored=$valueMode/$zerosShown"
    $analysis.TrimEnd()
    Write-Output 'VBA active-window value-display COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayFormulas -and $null -ne $originalDisplayZeros) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.DisplayFormulas = $originalDisplayFormulas
                $excel.ActiveWindow.DisplayZeros = $originalDisplayZeros
            }
        }
        catch {}
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
