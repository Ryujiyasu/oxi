# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-cell-drag-and-drop-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'cell-drag-and-drop-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalCellDragAndDrop = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalCellDragAndDrop = [bool]$excel.CellDragAndDrop

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'CellDragAndDropProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableCellDragAndDrop()
    Application.CellDragAndDrop = False
End Sub

Public Sub EnableCellDragAndDrop()
    Application.CellDragAndDrop = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!CellDragAndDropProbe.DisableCellDragAndDrop")
    $disabled = [bool]$excel.CellDragAndDrop
    if ($disabled) {
        throw 'Application.CellDragAndDrop COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!CellDragAndDropProbe.EnableCellDragAndDrop")
    $enabled = [bool]$excel.CellDragAndDrop
    if (-not $enabled) {
        throw 'Application.CellDragAndDrop COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.CellDragAndDrop = $originalCellDragAndDrop
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
        '[CellDragAndDropProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.CellDragAndDrop: changes Excel's process-global cell drag-and-drop interaction state"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'CellDragAndDropProbe' })[0]
    $dragAndDropFindings = @($probeModule.findings | Where-Object { $_.reason -like '*drag-and-drop interaction*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.CellDragAndDrop' -ne 2 -or
        $dragAndDropFindings.Count -ne 2) {
        throw 'Application.CellDragAndDrop JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.CellDragAndDrop JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM CellDragAndDrop states: original=$originalCellDragAndDrop disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.CellDragAndDrop COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalCellDragAndDrop) {
        try { $excel.CellDragAndDrop = $originalCellDragAndDrop } catch {}
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
