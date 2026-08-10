# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-document-action-task-pane-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'document-action-task-pane-probe.xlsm'
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
    $initialState = [bool]$excel.DisplayDocumentActionTaskPane

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DocumentActionTaskPaneProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function TryHideDocumentActionTaskPane() As Long
    On Error Resume Next
    Application.DisplayDocumentActionTaskPane = False
    TryHideDocumentActionTaskPane = Err.Number
End Function

Public Function TryShowDocumentActionTaskPane() As Long
    On Error Resume Next
    Application.DisplayDocumentActionTaskPane = True
    TryShowDocumentActionTaskPane = Err.Number
End Function
'@)

    $hideError = [int]$excel.Run("'$($workbook.Name)'!DocumentActionTaskPaneProbe.TryHideDocumentActionTaskPane")
    $stateAfterHide = [bool]$excel.DisplayDocumentActionTaskPane
    $showError = [int]$excel.Run("'$($workbook.Name)'!DocumentActionTaskPaneProbe.TryShowDocumentActionTaskPane")
    $stateAfterShow = [bool]$excel.DisplayDocumentActionTaskPane
    if ($hideError -ne 0 -or $stateAfterHide -or $showError -ne 1004 -or $stateAfterShow) {
        throw "DisplayDocumentActionTaskPane VBA behaviour mismatch: hideError=$hideError afterHide=$stateAfterHide showError=$showError afterShow=$stateAfterShow"
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
        '[DocumentActionTaskPaneProbe]',
        'procedures: 2, statements: 6, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayDocumentActionTaskPane: requests Excel's legacy document-action task-pane user interface; modern Excel accepts disabling but rejects enabling with error 1004"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DocumentActionTaskPaneProbe' })[0]
    $taskPaneFindings = @($probeModule.findings | Where-Object { $_.reason -like '*document-action task-pane*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayDocumentActionTaskPane' -ne 2 -or
        $taskPaneFindings.Count -ne 2) {
        throw 'DisplayDocumentActionTaskPane JSON analysis did not preserve its UI dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'DisplayDocumentActionTaskPane JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayDocumentActionTaskPane: initial=$initialState hideError=$hideError afterHide=$stateAfterHide showError=$showError afterShow=$stateAfterShow"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.DisplayDocumentActionTaskPane COM execution and analysis: PASS'
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
