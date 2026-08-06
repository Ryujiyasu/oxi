# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-bracket-expression-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'bracket-expression-probe.xlsm'
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
    $sheet.Name = 'Data'
    $sheet.Range('A1').Value2 = 21
    $sheet.Range('A2').Value2 = 2
    $sheet.Range('D1').Value2 = 'Amount'
    $sheet.Range('D2').Value2 = 10
    $sheet.Range('D3').Value2 = 5
    $table = $sheet.ListObjects.Add(1, $sheet.Range('D1:D3'), $null, 1)
    $table.Name = 'DataTable'

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'BracketProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function BracketCells() As Double
    BracketCells = [A1] + [A2]
End Function

Public Function BracketFormula() As Double
    BracketFormula = [SUM(A1:A2)]
End Function

Public Function BracketStructuredReference() As Double
    BracketStructuredReference = [SUM(DataTable[Amount])]
End Function
'@)

    $cellsValue = [double]$excel.Run("'$($workbook.Name)'!BracketProbe.BracketCells")
    $formulaValue = [double]$excel.Run("'$($workbook.Name)'!BracketProbe.BracketFormula")
    $structuredValue = [double]$excel.Run("'$($workbook.Name)'!BracketProbe.BracketStructuredReference")
    $storedSource = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
    if ($cellsValue -ne 23 -or $formulaValue -ne 23 -or $structuredValue -ne 15) {
        throw "Bracket COM execution mismatch: cells=$cellsValue formula=$formulaValue structured=$structuredValue"
    }
    foreach ($expectedSource in @(
        'BracketCells = [A1] + [A2]',
        'BracketFormula = [SUM(A1:A2)]',
        'BracketStructuredReference = [SUM(DataTable[Amount])]'
    )) {
        if (-not $storedSource.Contains($expectedSource)) {
            throw "VBE did not preserve bracket expression: $expectedSource"
        }
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
        '[BracketProbe]',
        'verdict: B (data transformation): evaluates an Excel name or formula',
        'procedures: 3, statements: 3',
        'unparsed: 0',
        'formula engine: required',
        'Evaluate: evaluates an Excel name or formula'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'BracketProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.Evaluate -ne 4 -or
        -not $probeModule.needs_formula_engine) {
        throw 'Bracket expression JSON analysis did not preserve Evaluate dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Bracket expression JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM values: [A1] + [A2] = $cellsValue; [SUM(A1:A2)] = $formulaValue; [SUM(DataTable[Amount])] = $structuredValue"
    Write-Output 'VBE stored all bracket expressions verbatim'
    $analysis.TrimEnd()
    Write-Output 'Excel bracket expression COM execution and analysis: PASS'
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
