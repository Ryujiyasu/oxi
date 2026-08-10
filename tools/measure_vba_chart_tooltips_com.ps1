# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-chart-tooltips-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'chart-tooltips-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalShowChartTipNames = $null
$originalShowChartTipValues = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalShowChartTipNames = [bool]$excel.ShowChartTipNames
    $originalShowChartTipValues = [bool]$excel.ShowChartTipValues

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ChartToolTipsProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub HideChartTipNames()
    Application.ShowChartTipNames = False
End Sub

Public Sub ShowChartTipNames()
    Application.ShowChartTipNames = True
End Sub

Public Sub HideChartTipValues()
    Application.ShowChartTipValues = False
End Sub

Public Sub ShowChartTipValues()
    Application.ShowChartTipValues = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!ChartToolTipsProbe.HideChartTipNames")
    $namesHidden = [bool]$excel.ShowChartTipNames
    if ($namesHidden) {
        throw 'Application.ShowChartTipNames COM hidden-state mismatch: expected False'
    }
    $excel.Run("'$($workbook.Name)'!ChartToolTipsProbe.ShowChartTipNames")
    $namesShown = [bool]$excel.ShowChartTipNames
    if (-not $namesShown) {
        throw 'Application.ShowChartTipNames COM shown-state mismatch: expected True'
    }

    $excel.Run("'$($workbook.Name)'!ChartToolTipsProbe.HideChartTipValues")
    $valuesHidden = [bool]$excel.ShowChartTipValues
    if ($valuesHidden) {
        throw 'Application.ShowChartTipValues COM hidden-state mismatch: expected False'
    }
    $excel.Run("'$($workbook.Name)'!ChartToolTipsProbe.ShowChartTipValues")
    $valuesShown = [bool]$excel.ShowChartTipValues
    if (-not $valuesShown) {
        throw 'Application.ShowChartTipValues COM shown-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.ShowChartTipNames = $originalShowChartTipNames
    $excel.ShowChartTipValues = $originalShowChartTipValues
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
        '[ChartToolTipsProbe]',
        'procedures: 4, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.ShowChartTipNames: shows or hides series names in Excel's process-global chart tooltip user interface",
        "Application.ShowChartTipValues: shows or hides values in Excel's process-global chart tooltip user interface"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'ChartToolTipsProbe' })[0]
    $nameFindings = @($probeModule.findings | Where-Object { $_.reason -like '*series names*' })
    $valueFindings = @($probeModule.findings | Where-Object { $_.reason -like '*values in Excel*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.ShowChartTipNames' -ne 2 -or
        $probeModule.api_names.'Application.ShowChartTipValues' -ne 2 -or
        $nameFindings.Count -ne 2 -or
        $valueFindings.Count -ne 2) {
        throw 'Chart-tooltip JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Chart-tooltip JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM chart tooltips: namesOriginal=$originalShowChartTipNames namesHidden=$namesHidden namesShown=$namesShown valuesOriginal=$originalShowChartTipValues valuesHidden=$valuesHidden valuesShown=$valuesShown"
    $analysis.TrimEnd()
    Write-Output 'VBA Application chart-tooltip COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        if ($null -ne $originalShowChartTipNames) {
            try { $excel.ShowChartTipNames = $originalShowChartTipNames } catch {}
        }
        if ($null -ne $originalShowChartTipValues) {
            try { $excel.ShowChartTipValues = $originalShowChartTipValues } catch {}
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
