# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-range-find-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'range-find-probe.xlsm'
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
    $sheet.Range('A1').Value2 = 1
    $sheet.Range('A2').Value2 = 42
    $sheet.Range('A3').Value2 = 42

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'RangeFindProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function FindValues() As String
    Dim found As Range
    Dim following As Range
    Set found = Sheet1.Range("A1:A3").Find(What:=42, After:=Sheet1.Range("A1"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    Set following = Sheet1.Range("A1:A3").FindNext(After:=found)
    FindValues = found.Address(False, False) & "|" & following.Address(False, False)
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!RangeFindProbe.FindValues")
    if ($actual -ne 'A2|A3') {
        throw "Range.Find COM execution mismatch: expected 'A2|A3', got '$actual'"
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
        '[RangeFindProbe]',
        'procedures: 1, statements: 5, max nesting: 0, unparsed: 0',
        'verdict: B',
        "Sheet1.Range.Find: uses Excel's stateful Range.Find settings; omitted options can inherit previous UI or VBA choices",
        "Sheet1.Range.FindNext: continues Excel's stateful preceding Range.Find operation"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'RangeFindProbe' })[0]
    $searchFindings = @($probeModule.findings | Where-Object { $_.reason -like '*stateful*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Sheet1.Range.Find' -ne 1 -or
        $probeModule.api_names.'Sheet1.Range.FindNext' -ne 1 -or
        $searchFindings.Count -ne 2) {
        throw 'Range.Find JSON analysis did not preserve stateful-search dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Range.Find JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Range.Find addresses: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Range.Find COM execution and analysis: PASS'
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
