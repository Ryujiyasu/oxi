# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-application-cursor-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'application-cursor-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalCursor = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalCursor = [int]$excel.Cursor
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ApplicationCursorProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub ShowWaitCursor()
    Application.Cursor = xlWait
End Sub

Public Sub RestoreDefaultCursor()
    Application.Cursor = xlDefault
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ApplicationCursorProbe.ShowWaitCursor")
    $waiting = [int]$excel.Cursor
    if ($waiting -ne 2) {
        throw "Application.Cursor COM wait-state mismatch: expected 2, got $waiting"
    }

    $excel.Run("'$($workbook.Name)'!ApplicationCursorProbe.RestoreDefaultCursor")
    $defaulted = [int]$excel.Cursor
    if ($defaulted -ne -4143) {
        throw "Application.Cursor COM default-state mismatch: expected -4143, got $defaulted"
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.Cursor = $originalCursor
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
        '[ApplicationCursorProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.Cursor: changes Excel's process-global mouse-pointer user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ApplicationCursorProbe' })[0]
    $cursorFindings = @($probeModule.findings | Where-Object { $_.reason -like '*mouse-pointer user interface*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.Cursor' -ne 2 -or
        $cursorFindings.Count -ne 2) {
        throw 'Application.Cursor JSON analysis did not preserve user-interface dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.Cursor JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Cursor states: original=$originalCursor waiting=$waiting defaulted=$defaulted"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.Cursor COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalCursor) {
        try { $excel.Cursor = $originalCursor } catch {}
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
