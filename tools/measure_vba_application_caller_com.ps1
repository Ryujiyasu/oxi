# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-application-caller-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'application-caller-probe.xlsm'
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
    $component.Name = 'CallerProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function CallerAddress() As String
    CallerAddress = Application.Caller.Address(False, False)
End Function
'@)

    $cell = $workbook.Worksheets.Item(1).Range('A1')
    $cell.Formula = '=CallerAddress()'
    $excel.Calculate()
    $actual = [string]$cell.Value2
    if ($actual -ne 'A1') {
        throw "Application.Caller COM execution mismatch: expected 'A1', got '$actual'"
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
        '[CallerProbe]',
        'procedures: 1, statements: 1, max nesting: 0, unparsed: 0',
        'Application.Caller.Address: reads the Excel invocation context; behavior depends on the calling cell or object'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'CallerProbe' })[0]
    $callerFindings = @($probeModule.findings | Where-Object { $_.reason -like '*calling cell or object*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.Caller.Address' -ne 1 -or
        $callerFindings.Count -ne 1) {
        throw 'Application.Caller JSON analysis did not preserve invocation-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.Caller JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Application.Caller address: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.Caller COM execution and analysis: PASS'
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
