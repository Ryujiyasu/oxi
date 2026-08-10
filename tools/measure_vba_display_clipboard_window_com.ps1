# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-display-clipboard-window-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'display-clipboard-window-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayClipboardWindow = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayClipboardWindow = [bool]$excel.DisplayClipboardWindow

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'DisplayClipboardWindowProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideClipboardWindow()
    Application.DisplayClipboardWindow = False
End Sub

Public Sub ShowClipboardWindow()
    Application.DisplayClipboardWindow = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!DisplayClipboardWindowProbe.HideClipboardWindow")
    $hidden = [bool]$excel.DisplayClipboardWindow
    if ($hidden) {
        throw 'Application.DisplayClipboardWindow COM hidden-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!DisplayClipboardWindowProbe.ShowClipboardWindow")
    $shownAfterRequest = [bool]$excel.DisplayClipboardWindow
    if ($shownAfterRequest) {
        throw 'Application.DisplayClipboardWindow modern-Excel behaviour changed: True request is no longer ignored'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.DisplayClipboardWindow = $originalDisplayClipboardWindow
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
        '[DisplayClipboardWindowProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.DisplayClipboardWindow: requests showing or hiding Excel's Office Clipboard task-pane user interface; modern Excel may ignore it"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'DisplayClipboardWindowProbe' })[0]
    $clipboardWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*Clipboard task-pane*' })
    $cutCopyFindings = @($probeModule.findings | Where-Object { $_.reason -like '*clipboard and cut/copy mode*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.DisplayClipboardWindow' -ne 2 -or
        $clipboardWindowFindings.Count -ne 2 -or
        $cutCopyFindings.Count -ne 0) {
        throw 'Application.DisplayClipboardWindow JSON analysis did not preserve its UI dependency'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.DisplayClipboardWindow JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayClipboardWindow states: original=$originalDisplayClipboardWindow hidden=$hidden shownAfterRequest=$shownAfterRequest"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.DisplayClipboardWindow COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayClipboardWindow) {
        try { $excel.DisplayClipboardWindow = $originalDisplayClipboardWindow } catch {}
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
