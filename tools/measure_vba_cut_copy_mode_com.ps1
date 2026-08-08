# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-cut-copy-mode-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'cut-copy-mode-probe.xlsm'
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

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'CutCopyModeProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ExerciseCutCopyMode() As String
    Sheet1.Range("A1").Value2 = "copied"
    Sheet1.Range("A1").Copy
    ExerciseCutCopyMode = CStr(CLng(Application.CutCopyMode))
    Application.CutCopyMode = False
    ExerciseCutCopyMode = ExerciseCutCopyMode & "|" & CStr(CLng(Application.CutCopyMode))
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!CutCopyModeProbe.ExerciseCutCopyMode")
    if ($actual -ne '1|0') {
        throw "CutCopyMode COM execution mismatch: expected '1|0', got '$actual'"
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
        '[CutCopyModeProbe]',
        'procedures: 1, statements: 5, max nesting: 0, unparsed: 0',
        'verdict: B',
        "Sheet1.Range.Copy: changes Excel's process-global clipboard and cut/copy mode",
        "Application.CutCopyMode: reads or changes Excel's process-global clipboard and cut/copy mode"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'CutCopyModeProbe' })[0]
    $clipboardFindings = @($probeModule.findings | Where-Object { $_.reason -like '*clipboard and cut/copy mode*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Sheet1.Range.Copy' -ne 1 -or
        $probeModule.api_names.'Application.CutCopyMode' -ne 3 -or
        $clipboardFindings.Count -ne 4) {
        throw 'CutCopyMode JSON analysis did not preserve global clipboard-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'CutCopyMode JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM CutCopyMode states: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA CutCopyMode COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        try { $excel.CutCopyMode = $false } catch {}
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
