# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-range-replace-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'range-replace-probe.xlsm'
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
    $sheet.Range('A1').Value2 = 'foo'
    $sheet.Range('A2').Value2 = 'bar'
    $sheet.Range('A3').Value2 = 'foo'

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'RangeReplaceProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ReplaceValues() As String
    Sheet1.Range("A1:A3").Replace What:="foo", Replacement:="baz", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    ReplaceValues = Sheet1.Range("A1").Value2 & "|" & Sheet1.Range("A2").Value2 & "|" & Sheet1.Range("A3").Value2
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!RangeReplaceProbe.ReplaceValues")
    if ($actual -ne 'baz|bar|baz') {
        throw "Range.Replace COM execution mismatch: expected 'baz|bar|baz', got '$actual'"
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
        '[RangeReplaceProbe]',
        'procedures: 1, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: B',
        "Sheet1.Range.Replace: uses Excel's stateful Range.Replace settings; omitted options can inherit previous UI or VBA choices"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'RangeReplaceProbe' })[0]
    $replaceFindings = @($probeModule.findings | Where-Object { $_.reason -like '*stateful Range.Replace*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Sheet1.Range.Replace' -ne 1 -or
        $replaceFindings.Count -ne 1) {
        throw 'Range.Replace JSON analysis did not preserve stateful-operation dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Range.Replace JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Range.Replace values: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Range.Replace COM execution and analysis: PASS'
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
