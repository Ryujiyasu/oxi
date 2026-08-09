# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-worksheet-page-breaks-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'worksheet-page-breaks-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$sheet = $null
$originalDisplayPageBreaks = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $originalDisplayPageBreaks = [bool]$sheet.DisplayPageBreaks

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'WorksheetPageBreaksProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HidePageBreaks()
    Sheet1.DisplayPageBreaks = False
End Sub

Public Sub ShowPageBreaksOnActiveSheet()
    Application.ActiveSheet.DisplayPageBreaks = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!WorksheetPageBreaksProbe.HidePageBreaks")
    $hidden = [bool]$sheet.DisplayPageBreaks
    if ($hidden) {
        throw 'Worksheet.DisplayPageBreaks COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!WorksheetPageBreaksProbe.ShowPageBreaksOnActiveSheet")
    $shown = [bool]$sheet.DisplayPageBreaks
    if (-not $shown) {
        throw 'Worksheet.DisplayPageBreaks COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $sheet.DisplayPageBreaks = $originalDisplayPageBreaks
    $workbook.Close($false)
    $workbook = $null
    $sheet = $null
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
        '[WorksheetPageBreaksProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        'Sheet1.DisplayPageBreaks: shows or hides automatic and manual page-break indicators on a worksheet',
        "Application.ActiveSheet.DisplayPageBreaks: reads Excel's active UI context"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'WorksheetPageBreaksProbe' })[0]
    $activeSheetFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $pageBreakFindings = @($probeModule.findings | Where-Object { $_.reason -like '*page-break indicators*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Sheet1.DisplayPageBreaks' -ne 1 -or
        $probeModule.api_names.'Application.ActiveSheet.DisplayPageBreaks' -ne 1 -or
        $activeSheetFindings.Count -ne 1 -or
        $pageBreakFindings.Count -ne 2) {
        throw 'Worksheet.DisplayPageBreaks JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Worksheet.DisplayPageBreaks JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayPageBreaks states: original=$originalDisplayPageBreaks hidden=$hidden shown=$shown"
    $analysis.TrimEnd()
    Write-Output 'VBA Worksheet.DisplayPageBreaks COM execution and analysis: PASS'
}
finally {
    if ($null -ne $sheet -and $null -ne $originalDisplayPageBreaks) {
        try { $sheet.DisplayPageBreaks = $originalDisplayPageBreaks } catch {}
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
