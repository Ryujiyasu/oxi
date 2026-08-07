# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-application-goto-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'application-goto-probe.xlsm'
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
    $component.Name = 'ApplicationGotoProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function GoToCell() As String
    Application.Goto Reference:=Sheet1.Range("C3"), Scroll:=True
    GoToCell = Application.ActiveCell.Address(False, False)
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!ApplicationGotoProbe.GoToCell")
    if ($actual -ne 'C3') {
        throw "Application.Goto COM execution mismatch: expected 'C3', got '$actual'"
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
        '[ApplicationGotoProbe]',
        'procedures: 1, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        'Application.Goto: activates and scrolls to an Excel range or object',
        "Application.ActiveCell.Address: reads Excel's active UI context"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ApplicationGotoProbe' })[0]
    $gotoFindings = @($probeModule.findings | Where-Object { $_.reason -like '*activates and scrolls*' })
    $contextFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.Goto' -ne 1 -or
        $probeModule.api_names.'Application.ActiveCell.Address' -ne 1 -or
        $gotoFindings.Count -ne 1 -or
        $contextFindings.Count -ne 1) {
        throw 'Application.Goto JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.Goto JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Application.Goto active cell: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.Goto COM execution and analysis: PASS'
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
