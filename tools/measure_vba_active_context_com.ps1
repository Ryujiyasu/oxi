# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-context-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-context-probe.xlsm'
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
    $sheet = $workbook.Worksheets.Item(1)

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveContextProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ReadActiveContext() As String
    Sheet1.Activate
    Sheet1.Range("B2").Select
    ReadActiveContext = Application.ActiveWorkbook.Name & "|" & Application.ActiveSheet.Name & "|" & Application.ActiveCell.Address(False, False) & "|" & Application.Selection.Address(False, False)
End Function
'@)

    $expected = "$($workbook.Name)|$($sheet.Name)|B2|B2"
    $actual = [string]$excel.Run("'$($workbook.Name)'!ActiveContextProbe.ReadActiveContext")
    if ($actual -ne $expected) {
        throw "Active-context COM execution mismatch: expected '$expected', got '$actual'"
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
        '[ActiveContextProbe]',
        'procedures: 1, statements: 3, max nesting: 0, unparsed: 0',
        'verdict: B',
        "reads Excel's active UI context; result depends on the current workbook, sheet, cell, or selection"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveContextProbe' })[0]
    $contextFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ActiveWorkbook.Name' -ne 1 -or
        $probeModule.api_names.'Application.ActiveSheet.Name' -ne 1 -or
        $probeModule.api_names.'Application.ActiveCell.Address' -ne 1 -or
        $probeModule.api_names.'Application.Selection.Address' -ne 1 -or
        $contextFindings.Count -ne 4) {
        throw 'Active-context JSON analysis did not preserve current UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Active-context JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM active context: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA active-context COM execution and analysis: PASS'
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
