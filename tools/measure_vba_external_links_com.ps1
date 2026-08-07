# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-external-links-" + [guid]::NewGuid().ToString('N'))
$sourcePath = Join-Path $probeRoot 'external-link-source.xlsx'
$workbookPath = Join-Path $probeRoot 'external-links-probe.xlsm'
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
    $externalFormula = "='$probeRoot\[external-link-source.xlsx]Sheet1'!`$A`$1"
    $workbook.Worksheets.Item(1).Range('A1').Formula = $externalFormula

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ExternalLinksProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function RefreshExternalLink() As Double
    Dim links As Variant
    links = ThisWorkbook.LinkSources(xlExcelLinks)
    ThisWorkbook.UpdateLink Name:=links(1), Type:=xlExcelLinks
    RefreshExternalLink = ThisWorkbook.Worksheets(1).Range("A1").Value2
End Function
'@)

    $actual = [double]$excel.Run("'$($workbook.Name)'!ExternalLinksProbe.RefreshExternalLink")
    if ($actual -ne 42) {
        throw "External-link COM execution mismatch: expected 42, got $actual"
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
        '[ExternalLinksProbe]',
        'procedures: 1, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: C',
        'ThisWorkbook.LinkSources: enumerates external workbook links',
        'ThisWorkbook.UpdateLink: refreshes data from an external workbook link'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ExternalLinksProbe' })[0]
    $linkFindings = @($probeModule.findings | Where-Object { $_.reason -like '*external workbook link*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ThisWorkbook.LinkSources' -ne 1 -or
        $probeModule.api_names.'ThisWorkbook.UpdateLink' -ne 1 -or
        $linkFindings.Count -ne 2) {
        throw 'External-link JSON analysis did not preserve workbook-link dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'External-link JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM refreshed external-link value: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA external-link COM execution and analysis: PASS'
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
