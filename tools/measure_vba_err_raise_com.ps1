# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-err-raise-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'err-raise-probe.xlsm'
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
    $component.Name = 'ErrRaiseProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function CaptureRaisedError() As String
    On Error GoTo Handler
    Err.Raise 513, "OxiProbe", "boom"
    CaptureRaisedError = "not raised"
    Exit Function
Handler:
    CaptureRaisedError = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!ErrRaiseProbe.CaptureRaisedError")
    if ($actual -ne '513|OxiProbe|boom') {
        throw "Err.Raise COM execution mismatch: expected '513|OxiProbe|boom', got '$actual'"
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
        '[ErrRaiseProbe]',
        'procedures: 1, statements: 6, max nesting: 0, unparsed: 0',
        'Err.Raise: raises a VBA runtime error whose number, source, and description are observable'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ErrRaiseProbe' })[0]
    $raiseFindings = @($probeModule.findings | Where-Object { $_.reason -like '*number, source, and description*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Err.Raise' -ne 1 -or
        $raiseFindings.Count -ne 1) {
        throw 'Err.Raise JSON analysis did not preserve observable error semantics'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Err.Raise JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM captured Err.Raise: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Err.Raise COM execution and analysis: PASS'
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
