# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-search-formats-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'search-formats-probe.xlsm'
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
    $component.Name = 'SearchFormatsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ExerciseSearchFormats() As Long
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear
    Application.FindFormat.Font.Bold = True
    Application.ReplaceFormat.Font.Italic = True
    ExerciseSearchFormats = CLng(Application.FindFormat.Font.Bold) + CLng(Application.ReplaceFormat.Font.Italic)
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear
End Function
'@)

    $actual = [long]$excel.Run("'$($workbook.Name)'!SearchFormatsProbe.ExerciseSearchFormats")
    if ($actual -ne -2) {
        throw "Search-format COM execution mismatch: expected -2, got $actual"
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
        '[SearchFormatsProbe]',
        'procedures: 1, statements: 7, max nesting: 0, unparsed: 0',
        'verdict: A',
        "Application.FindFormat.Clear: reads or changes Excel's process-global find-format criteria",
        "Application.ReplaceFormat.Clear: reads or changes Excel's process-global replace-format criteria"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'SearchFormatsProbe' })[0]
    $findFormatFindings = @($probeModule.findings | Where-Object { $_.reason -like '*find-format criteria*' })
    $replaceFormatFindings = @($probeModule.findings | Where-Object { $_.reason -like '*replace-format criteria*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.FindFormat.Clear' -ne 2 -or
        $probeModule.api_names.'Application.ReplaceFormat.Clear' -ne 2 -or
        $probeModule.api_names.'Application.FindFormat.Font.Bold' -ne 2 -or
        $probeModule.api_names.'Application.ReplaceFormat.Font.Italic' -ne 2 -or
        $findFormatFindings.Count -ne 4 -or
        $replaceFormatFindings.Count -ne 4) {
        throw 'Search-format JSON analysis did not preserve global format-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Search-format JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM search-format Boolean sum: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA search-format COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        try { $excel.FindFormat.Clear() } catch {}
        try { $excel.ReplaceFormat.Clear() } catch {}
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
