# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-identifier-suffix-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'identifier-suffix-probe.xlsm'
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
    $component.Name = 'SuffixProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function VariantLeftPreservesNull() As Boolean
    VariantLeftPreservesNull = IsNull(Left(Null, 1))
End Function

Public Function StringLeftNullError() As Long
    On Error GoTo Handler
    Dim result As String
    result = Left$(Null, 1)
    StringLeftNullError = 0
    Exit Function
Handler:
    StringLeftNullError = Err.Number
End Function
'@)

    $variantPreservesNull = [bool]$excel.Run("'$($workbook.Name)'!SuffixProbe.VariantLeftPreservesNull")
    $stringError = [long]$excel.Run("'$($workbook.Name)'!SuffixProbe.StringLeftNullError")
    $storedSource = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
    if (-not $variantPreservesNull -or $stringError -ne 94) {
        throw "Identifier suffix COM mismatch: variant=$variantPreservesNull stringError=$stringError"
    }
    if (-not $storedSource.Contains('Left(Null, 1)') -or
        -not $storedSource.Contains('Left$(Null, 1)')) {
        throw 'VBE did not preserve the suffixed and unsuffixed calls'
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
        '[SuffixProbe]',
        'procedures: 2',
        'unparsed: 0'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'SuffixProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.Left -ne 2) {
        throw 'Identifier suffix JSON analysis did not retain the undecorated API name'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Identifier suffix JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM values: Left(Null, 1) preserves Null=$variantPreservesNull; Left`$(Null, 1) error=$stringError"
    Write-Output 'VBE stored the suffixed and unsuffixed calls distinctly'
    $analysis.TrimEnd()
    Write-Output 'VBA identifier suffix COM execution and analysis: PASS'
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
