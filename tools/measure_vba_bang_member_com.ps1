# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-bang-member-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'bang-member-probe.xlsm'
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
    $component.Name = 'BangProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function BangValue() As Long
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values.Add "Answer", 42
    values.Add "Display Name", 8
    BangValue = values!Answer + values![Display Name]
End Function

Public Function WithBangValue() As Long
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values.Add "Answer", 42
    values.Add "Display Name", 8
    With values
        WithBangValue = !Answer + ![Display Name]
    End With
End Function
'@)

    $actual = [long]$excel.Run("'$($workbook.Name)'!BangProbe.BangValue")
    $withActual = [long]$excel.Run("'$($workbook.Name)'!BangProbe.WithBangValue")
    $storedSource = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
    if ($actual -ne 50 -or $withActual -ne 50) {
        throw "Bang member COM execution mismatch: direct=$actual with=$withActual"
    }
    if (-not $storedSource.Contains('values!Answer + values![Display Name]')) {
        throw 'VBE did not preserve bang member access'
    }
    if (-not $storedSource.Contains('WithBangValue = !Answer + ![Display Name]')) {
        throw 'VBE did not preserve leading bang access inside With'
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
        '[BangProbe]',
        'verdict: C (out of scope: reaches outside Excel)',
        'procedures: 2, statements: 11',
        'unparsed: 0',
        'CreateObject: late-binds an external COM object'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'BangProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'values.Answer' -ne 1 -or
        $probeModule.api_names.'values.Display Name' -ne 1) {
        throw 'Bang member JSON analysis did not preserve default-member names'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Bang member JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM values: direct bang=$actual; leading bang inside With=$withActual"
    Write-Output 'VBE stored direct and With-relative bang forms verbatim'
    $analysis.TrimEnd()
    Write-Output 'VBA bang member COM execution and analysis: PASS'
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
