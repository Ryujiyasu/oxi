# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-operation-options-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'operation-options-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalPasteOptions = $null
$originalInsertOptions = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalPasteOptions = [bool]$excel.DisplayPasteOptions
    $originalInsertOptions = [bool]$excel.DisplayInsertOptions

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'OperationOptionsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HidePasteOptions()
    Application.DisplayPasteOptions = False
End Sub

Public Sub ShowPasteOptions()
    Application.DisplayPasteOptions = True
End Sub

Public Sub HideInsertOptions()
    Application.DisplayInsertOptions = False
End Sub

Public Sub ShowInsertOptions()
    Application.DisplayInsertOptions = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!OperationOptionsProbe.HidePasteOptions")
    $pasteHidden = [bool]$excel.DisplayPasteOptions
    if ($pasteHidden) {
        throw 'Application.DisplayPasteOptions COM hidden-state mismatch: expected False'
    }
    $excel.Run("'$($workbook.Name)'!OperationOptionsProbe.ShowPasteOptions")
    $pasteShown = [bool]$excel.DisplayPasteOptions
    if (-not $pasteShown) {
        throw 'Application.DisplayPasteOptions COM shown-state mismatch: expected True'
    }

    $excel.Run("'$($workbook.Name)'!OperationOptionsProbe.HideInsertOptions")
    $insertHidden = [bool]$excel.DisplayInsertOptions
    if ($insertHidden) {
        throw 'Application.DisplayInsertOptions COM hidden-state mismatch: expected False'
    }
    $excel.Run("'$($workbook.Name)'!OperationOptionsProbe.ShowInsertOptions")
    $insertShown = [bool]$excel.DisplayInsertOptions
    if (-not $insertShown) {
        throw 'Application.DisplayInsertOptions COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.DisplayPasteOptions = $originalPasteOptions
    $excel.DisplayInsertOptions = $originalInsertOptions
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
        '[OperationOptionsProbe]',
        'procedures: 4, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayPasteOptions: shows or hides Excel's process-global paste-options user interface",
        "Application.DisplayInsertOptions: shows or hides Excel's process-global insert-options user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'OperationOptionsProbe' })[0]
    $pasteFindings = @($probeModule.findings | Where-Object { $_.reason -like '*paste-options*' })
    $insertFindings = @($probeModule.findings | Where-Object { $_.reason -like '*insert-options*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayPasteOptions' -ne 2 -or
        $probeModule.api_names.'Application.DisplayInsertOptions' -ne 2 -or
        $pasteFindings.Count -ne 2 -or
        $insertFindings.Count -ne 2) {
        throw 'Operation-options JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Operation-options JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM operation options: pasteOriginal=$originalPasteOptions pasteHidden=$pasteHidden pasteShown=$pasteShown insertOriginal=$originalInsertOptions insertHidden=$insertHidden insertShown=$insertShown"
    $analysis.TrimEnd()
    Write-Output 'VBA Application operation-options COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        if ($null -ne $originalPasteOptions) {
            try { $excel.DisplayPasteOptions = $originalPasteOptions } catch {}
        }
        if ($null -ne $originalInsertOptions) {
            try { $excel.DisplayInsertOptions = $originalInsertOptions } catch {}
        }
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
