# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-enable-auto-complete-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'enable-auto-complete-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalEnableAutoComplete = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalEnableAutoComplete = [bool]$excel.EnableAutoComplete

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'EnableAutoCompleteProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableCellAutoComplete()
    Application.EnableAutoComplete = False
End Sub

Public Sub EnableCellAutoComplete()
    Application.EnableAutoComplete = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!EnableAutoCompleteProbe.DisableCellAutoComplete")
    $disabled = [bool]$excel.EnableAutoComplete
    if ($disabled) {
        throw 'Application.EnableAutoComplete COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!EnableAutoCompleteProbe.EnableCellAutoComplete")
    $enabled = [bool]$excel.EnableAutoComplete
    if (-not $enabled) {
        throw 'Application.EnableAutoComplete COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.EnableAutoComplete = $originalEnableAutoComplete
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
        '[EnableAutoCompleteProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.EnableAutoComplete: changes Excel's process-global cell-entry AutoComplete user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'EnableAutoCompleteProbe' })[0]
    $autoCompleteFindings = @($probeModule.findings | Where-Object { $_.reason -like '*cell-entry AutoComplete*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.EnableAutoComplete' -ne 2 -or
        $autoCompleteFindings.Count -ne 2) {
        throw 'Application.EnableAutoComplete JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.EnableAutoComplete JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM EnableAutoComplete states: original=$originalEnableAutoComplete disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.EnableAutoComplete COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalEnableAutoComplete) {
        try { $excel.EnableAutoComplete = $originalEnableAutoComplete } catch {}
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
