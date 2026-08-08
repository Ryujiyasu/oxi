# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-active-window-rtl-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'active-window-rtl-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalDisplayRightToLeft = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalDisplayRightToLeft = [bool]$excel.ActiveWindow.DisplayRightToLeft

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ActiveWindowRtlProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub EnableRightToLeft()
    ActiveWindow.DisplayRightToLeft = True
End Sub

Public Sub DisableRightToLeft()
    Application.ActiveWindow.DisplayRightToLeft = False
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ActiveWindowRtlProbe.EnableRightToLeft")
    $enabled = [bool]$excel.ActiveWindow.DisplayRightToLeft
    if (-not $enabled) {
        throw 'ActiveWindow.DisplayRightToLeft COM enabled-state mismatch: expected True'
    }

    $excel.Run("'$($workbook.Name)'!ActiveWindowRtlProbe.DisableRightToLeft")
    $disabled = [bool]$excel.ActiveWindow.DisplayRightToLeft
    if ($disabled) {
        throw 'ActiveWindow.DisplayRightToLeft COM disabled-state mismatch: expected False'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ActiveWindow.DisplayRightToLeft = $originalDisplayRightToLeft
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
        '[ActiveWindowRtlProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        'verdict: D',
        "ActiveWindow.DisplayRightToLeft: reads Excel's active UI context",
        'ActiveWindow.DisplayRightToLeft: changes the worksheet display direction of an Excel window',
        'Application.ActiveWindow.DisplayRightToLeft: changes the worksheet display direction of an Excel window'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ActiveWindowRtlProbe' })[0]
    $activeWindowFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active UI context*' })
    $directionFindings = @($probeModule.findings | Where-Object { $_.reason -like '*worksheet display direction*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'ActiveWindow.DisplayRightToLeft' -ne 1 -or
        $probeModule.api_names.'Application.ActiveWindow.DisplayRightToLeft' -ne 1 -or
        $activeWindowFindings.Count -ne 2 -or
        $directionFindings.Count -ne 2) {
        throw 'ActiveWindow.DisplayRightToLeft JSON analysis did not preserve UI-context dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'ActiveWindow.DisplayRightToLeft JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM DisplayRightToLeft states: original=$originalDisplayRightToLeft enabled=$enabled disabled=$disabled"
    $analysis.TrimEnd()
    Write-Output 'VBA ActiveWindow.DisplayRightToLeft COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalDisplayRightToLeft) {
        try {
            if ($null -ne $excel.ActiveWindow) {
                $excel.ActiveWindow.DisplayRightToLeft = $originalDisplayRightToLeft
            }
        }
        catch {}
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
