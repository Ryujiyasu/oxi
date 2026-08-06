# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-with-member-analysis-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'with-member-analysis-probe.xlsm'
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
    $component.Name = 'WithMemberProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub FormatWith()
    With ThisWorkbook.Worksheets(1).Range("A1")
        .Value2 = 7
        .Font.Bold = True
    End With
End Sub

Public Function ReadWith() As Long
    With ThisWorkbook.Worksheets(1).Range("A1")
        ReadWith = .Value2
    End With
End Function
'@)

    $excel.Run("'$($workbook.Name)'!WithMemberProbe.FormatWith")
    $actual = [long]$excel.Run("'$($workbook.Name)'!WithMemberProbe.ReadWith")
    $bold = [bool]$workbook.Worksheets.Item(1).Range('A1').Font.Bold
    if ($actual -ne 7 -or -not $bold) {
        throw "With member COM execution mismatch: value=$actual bold=$bold"
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
    foreach ($expected in @(
        '[WithMemberProbe]',
        'verdict: A (report generation): sets fonts',
        'procedures: 2, statements: 5, max nesting: 1, unparsed: 0',
        'ThisWorkbook.Worksheets.Range.Font.Bold: sets fonts'
    )) {
        if (-not $analysis.Contains($expected)) {
            throw "Expected analysis fragment not found: $expected`n$analysis"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'WithMemberProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ThisWorkbook.Worksheets.Range.Value2' -ne 2 -or
        $probeModule.api_names.'ThisWorkbook.Worksheets.Range.Font.Bold' -ne 1) {
        throw 'With member JSON analysis did not resolve relative member chains'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'With member JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM values: A1.Value2=$actual; A1.Font.Bold=$bold"
    $analysis.TrimEnd()
    Write-Output 'VBA With-relative member COM execution and analysis: PASS'
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
