# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-convert-formula-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'convert-formula-probe.xlsm'
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
    $component.Name = 'ConvertFormulaProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ConvertedAddress() As String
    ConvertedAddress = Application.ConvertFormula("=R2C2", xlR1C1, xlA1, xlAbsolute)
End Function
'@)

    $expected = [string]$excel.ConvertFormula('=R2C2', -4150, 1, 1)
    $actual = [string]$excel.Run("'$($workbook.Name)'!ConvertFormulaProbe.ConvertedAddress")
    if ($actual -ne $expected -or $actual -ne '=$B$2') {
        throw "ConvertFormula COM execution mismatch: expected '=$B$2' / '$expected', got '$actual'"
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
        '[ConvertFormulaProbe]',
        'verdict: B (data transformation): converts formula references between A1 and R1C1 notation',
        'procedures: 1, statements: 1, max nesting: 0, unparsed: 0',
        'formula engine: required',
        'Application.ConvertFormula: converts formula references between A1 and R1C1 notation'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ConvertFormulaProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        -not $probeModule.needs_formula_engine -or
        $probeModule.api_names.'Application.ConvertFormula' -ne 1) {
        throw 'ConvertFormula JSON analysis did not preserve the formula conversion dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ConvertFormula JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM ConvertFormula result: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA ConvertFormula COM execution and analysis: PASS'
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
