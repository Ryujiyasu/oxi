# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-workbooks-open-" + [guid]::NewGuid().ToString('N'))
$sourcePath = Join-Path $probeRoot 'external-source.xlsx'
$workbookPath = Join-Path $probeRoot 'workbooks-open-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$sourceWorkbook = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $sourceWorkbook = $excel.Workbooks.Add()
    $sourceWorkbook.Worksheets.Item(1).Range('A1').Value2 = 42
    $sourceWorkbook.SaveAs($sourcePath, 51)
    $sourceWorkbook.Close($false)
    $sourceWorkbook = $null

    $workbook = $excel.Workbooks.Add()
    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ExternalWorkbookProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ReadExternalWorkbook(ByVal path As String) As Double
    Dim opened As Workbook
    Set opened = Application.Workbooks.Open(Filename:=path, ReadOnly:=True)
    ReadExternalWorkbook = opened.Worksheets(1).Range("A1").Value2
    opened.Close SaveChanges:=False
End Function
'@)

    $actual = [double]$excel.Run("'$($workbook.Name)'!ExternalWorkbookProbe.ReadExternalWorkbook", $sourcePath)
    if ($actual -ne 42) {
        throw "Workbooks.Open COM execution mismatch: expected 42, got $actual"
    }
    if ($excel.Workbooks.Count -ne 1) {
        throw "Workbooks.Open COM cleanup mismatch: expected one open probe workbook, got $($excel.Workbooks.Count)"
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
        '[ExternalWorkbookProbe]',
        'procedures: 1, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: C',
        'Application.Workbooks.Open: opens an external workbook from a path'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ExternalWorkbookProbe' })[0]
    $openFindings = @($probeModule.findings | Where-Object { $_.reason -like '*external workbook from a path*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.Workbooks.Open' -ne 1 -or
        $openFindings.Count -ne 1) {
        throw 'Workbooks.Open JSON analysis did not preserve external-workbook dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Workbooks.Open JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM external workbook value: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Workbooks.Open COM execution and analysis: PASS'
}
finally {
    if ($null -ne $sourceWorkbook) {
        try { $sourceWorkbook.Close($false) } catch {}
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
