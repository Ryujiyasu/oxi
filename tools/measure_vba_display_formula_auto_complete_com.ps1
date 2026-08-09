# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-formula-auto-complete-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-formula-auto-complete-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayFormulaAutoComplete = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayFormulaAutoComplete = [bool]$excel.DisplayFormulaAutoComplete

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayFormulaAutoCompleteProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableFormulaAutoComplete()
    Application.DisplayFormulaAutoComplete = False
End Sub

Public Sub EnableFormulaAutoComplete()
    Application.DisplayFormulaAutoComplete = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayFormulaAutoCompleteProbe.DisableFormulaAutoComplete")
    $disabled = [bool]$excel.DisplayFormulaAutoComplete
    if ($disabled) {
        throw 'Application.DisplayFormulaAutoComplete COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayFormulaAutoCompleteProbe.EnableFormulaAutoComplete")
    $enabled = [bool]$excel.DisplayFormulaAutoComplete
    if (-not $enabled) {
        throw 'Application.DisplayFormulaAutoComplete COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.DisplayFormulaAutoComplete = $originalDisplayFormulaAutoComplete
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
        '[DisplayFormulaAutoCompleteProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        'formula engine: not detected',
        "Application.DisplayFormulaAutoComplete: changes Excel's process-global formula-entry AutoComplete user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayFormulaAutoCompleteProbe' })[0]
    $autoCompleteFindings = @($probeModule.findings | Where-Object { $_.reason -like '*formula-entry AutoComplete*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.needs_formula_engine -or
        $probeModule.api_names.'Application.DisplayFormulaAutoComplete' -ne 2 -or
        $autoCompleteFindings.Count -ne 2) {
        throw 'Application.DisplayFormulaAutoComplete JSON analysis did not preserve UI-only dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.DisplayFormulaAutoComplete JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayFormulaAutoComplete states: original=$originalDisplayFormulaAutoComplete disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.DisplayFormulaAutoComplete COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayFormulaAutoComplete) {
        try { $excel.DisplayFormulaAutoComplete = $originalDisplayFormulaAutoComplete } catch {}
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
