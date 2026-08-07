# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-international-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'international-probe.xlsm'
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
    $component.Name = 'InternationalProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function LocaleSeparators() As String
    LocaleSeparators = Application.International(xlDecimalSeparator) & "|" & Application.International(xlListSeparator)
End Function
'@)

    $decimalSeparator = [string]$excel.International(3)
    $listSeparator = [string]$excel.International(5)
    $expected = "$decimalSeparator|$listSeparator"
    $actual = [string]$excel.Run("'$($workbook.Name)'!InternationalProbe.LocaleSeparators")
    if ($actual -ne $expected) {
        throw "Application.International COM execution mismatch: expected '$expected', got '$actual'"
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
        '[InternationalProbe]',
        'procedures: 1, statements: 1, max nesting: 0, unparsed: 0',
        'Application.International: reads Excel locale settings; behavior can vary by machine'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'InternationalProbe' })[0]
    $localeFindings = @($probeModule.findings | Where-Object { $_.reason -like '*locale settings*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.International' -ne 2 -or
        $localeFindings.Count -ne 2) {
        throw 'Application.International JSON analysis did not preserve locale dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.International JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM locale separators: decimal='$decimalSeparator', list='$listSeparator'"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.International COM execution and analysis: PASS'
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
