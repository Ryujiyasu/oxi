# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-file-dialog-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'file-dialog-probe.xlsm'
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
    $component.Name = 'FileDialogProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function ConfigureFileDialog() As String
    Dim picker As Object
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    picker.Title = "Oxi File Dialog Probe"
    ConfigureFileDialog = picker.Title
End Function
'@)

    $actual = [string]$excel.Run("'$($workbook.Name)'!FileDialogProbe.ConfigureFileDialog")
    if ($actual -ne 'Oxi File Dialog Probe') {
        throw "FileDialog COM execution mismatch: expected 'Oxi File Dialog Probe', got '$actual'"
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
        '[FileDialogProbe]',
        'verdict: D (out of scope: has its own user interface): opens an Office file-selection user interface',
        'procedures: 1, statements: 4, max nesting: 0, unparsed: 0',
        '[D] Application.FileDialog: opens an Office file-selection user interface'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'FileDialogProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.class -ne 'D' -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.FileDialog' -ne 1) {
        throw 'FileDialog JSON analysis did not preserve the user-interface dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'FileDialog JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM FileDialog title: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA FileDialog COM execution and analysis: PASS'
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
