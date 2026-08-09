# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-comment-indicator-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-comment-indicator-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayCommentIndicator = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayCommentIndicator = [int]$excel.DisplayCommentIndicator

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayCommentIndicatorProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideNoteIndicators()
    Application.DisplayCommentIndicator = xlNoIndicator
End Sub

Public Sub ShowNoteIndicators()
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
End Sub

Public Sub ShowNotesAndIndicators()
    Application.DisplayCommentIndicator = xlCommentAndIndicator
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayCommentIndicatorProbe.HideNoteIndicators")
    $hidden = [int]$excel.DisplayCommentIndicator
    if ($hidden -ne 0) {
        throw "Application.DisplayCommentIndicator COM hidden mismatch: expected 0, got $hidden"
    }

    $excel.Run("'$($workbook.Name)'!DisplayCommentIndicatorProbe.ShowNoteIndicators")
    $indicatorOnly = [int]$excel.DisplayCommentIndicator
    if ($indicatorOnly -ne -1) {
        throw "Application.DisplayCommentIndicator COM indicator-only mismatch: expected -1, got $indicatorOnly"
    }

    $excel.Run("'$($workbook.Name)'!DisplayCommentIndicatorProbe.ShowNotesAndIndicators")
    $notesAndIndicators = [int]$excel.DisplayCommentIndicator
    if ($notesAndIndicators -ne 1) {
        throw "Application.DisplayCommentIndicator COM notes-and-indicators mismatch: expected 1, got $notesAndIndicators"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.DisplayCommentIndicator = $originalDisplayCommentIndicator
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
        '[DisplayCommentIndicatorProbe]',
        'procedures: 3, statements: 3, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayCommentIndicator: changes Excel's process-global cell-note indicator and comment display user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayCommentIndicatorProbe' })[0]
    $commentDisplayFindings = @($probeModule.findings | Where-Object { $_.reason -like '*cell-note indicator*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayCommentIndicator' -ne 3 -or
        $commentDisplayFindings.Count -ne 3) {
        throw 'Application.DisplayCommentIndicator JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.DisplayCommentIndicator JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayCommentIndicator values: original=$originalDisplayCommentIndicator hidden=$hidden indicatorOnly=$indicatorOnly notesAndIndicators=$notesAndIndicators"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.DisplayCommentIndicator COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayCommentIndicator) {
        try { $excel.DisplayCommentIndicator = $originalDisplayCommentIndicator } catch {}
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
