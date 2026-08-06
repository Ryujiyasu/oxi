# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-input-function-" + [guid]::NewGuid().ToString('N'))
$inputPath = Join-Path $probeRoot 'input.txt'
$workbookPath = Join-Path $probeRoot 'input-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    [IO.File]::WriteAllText($inputPath, 'ABCD', [Text.Encoding]::ASCII)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'InputProbe'
    $module.CodeModule.AddFromString(@'
Option Explicit

Public Function ReadPrefix(ByVal path As String) As String
    Dim handle As Integer
    handle = FreeFile
    Open path For Input As #handle
    ReadPrefix = Input$(2, #handle)
    Close #handle
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!InputProbe.ReadPrefix", $inputPath)
    if ($actual -ne 'AB') {
        throw "Input$ COM execution mismatch: expected AB, got [$actual]"
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
        '[InputProbe]',
        'verdict: C (out of scope: reaches outside Excel)',
        'procedures: 1, statements: 5',
        'unparsed: 0',
        'Input: reads bytes or characters from an external file'
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
    $inputModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'InputProbe' })[0]
    if ($null -eq $inputModule -or $inputModule.metrics.unparsed -ne 0 -or $inputModule.api_names.Input -ne 1) {
        throw 'Input$ JSON analysis did not preserve the hash file-number argument'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Input$ JSON inventory reported unexpected errors'
    }

    $analysis.TrimEnd()
    Write-Output 'Input$(count, #fileNumber) COM execution and analysis: PASS'
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
