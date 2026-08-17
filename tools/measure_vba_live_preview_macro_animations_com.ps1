# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-live-preview-macro-animations-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'live-preview-macro-animations-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalLivePreview = $null
$originalMacroAnimations = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalLivePreview = [bool]$excel.EnableLivePreview
    $originalMacroAnimations = [bool]$excel.EnableMacroAnimations

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'AnimatedUiProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableLivePreview()
    Application.EnableLivePreview = False
End Sub

Public Sub EnableLivePreview()
    Application.EnableLivePreview = True
End Sub

Public Sub DisableMacroAnimations()
    Application.EnableMacroAnimations = False
End Sub

Public Sub EnableMacroAnimations()
    Application.EnableMacroAnimations = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!AnimatedUiProbe.DisableLivePreview")
    $liveDisabled = [bool]$excel.EnableLivePreview
    $excel.Run("'$($workbook.Name)'!AnimatedUiProbe.EnableLivePreview")
    $liveEnabled = [bool]$excel.EnableLivePreview
    $excel.Run("'$($workbook.Name)'!AnimatedUiProbe.DisableMacroAnimations")
    $macroDisabled = [bool]$excel.EnableMacroAnimations
    $excel.Run("'$($workbook.Name)'!AnimatedUiProbe.EnableMacroAnimations")
    $macroEnabled = [bool]$excel.EnableMacroAnimations
    if ($liveDisabled -or -not $liveEnabled -or $macroDisabled -or -not $macroEnabled) {
        throw 'Excel COM animated-UI state did not follow False/True assignments'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.EnableLivePreview = $originalLivePreview
    $excel.EnableMacroAnimations = $originalMacroAnimations
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    & cargo build -q -p oxidocs-cli --manifest-path (Join-Path $repoRoot 'Cargo.toml')
    if ($LASTEXITCODE -ne 0) { throw "oxidocs-cli build failed with exit code $LASTEXITCODE" }

    $analysis = (& $cliPath vba-analyze $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "vba-analyze failed with exit code $LASTEXITCODE" }
    foreach ($expectedFragment in @(
        '[AnimatedUiProbe]',
        'procedures: 4, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.EnableLivePreview: changes Excel's process-global live-preview user interface",
        'Application.EnableMacroAnimations: changes whether Excel shows user-interface animations while a macro runs'
    )) {
        if (-not $analysis.Contains($expectedFragment)) {
            throw "Expected analysis fragment not found: $expectedFragment`n$analysis"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "vba-inventory-json failed with exit code $LASTEXITCODE" }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'AnimatedUiProbe' })[0]
    $liveFindings = @($probeModule.findings | Where-Object { $_.reason -like '*live-preview user interface*' })
    $macroFindings = @($probeModule.findings | Where-Object { $_.reason -like '*while a macro runs*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.EnableLivePreview' -ne 2 -or
        $probeModule.api_names.'Application.EnableMacroAnimations' -ne 2 -or
        $liveFindings.Count -ne 2 -or $macroFindings.Count -ne 2) {
        throw 'Animated-UI JSON analysis did not preserve both UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) { throw 'Animated-UI JSON inventory reported unexpected errors' }

    Write-Output "Excel COM animated UI: live original=$originalLivePreview false=$liveDisabled true=$liveEnabled; macro original=$originalMacroAnimations false=$macroDisabled true=$macroEnabled"
    $analysis.TrimEnd()
    Write-Output 'VBA live-preview and macro-animation COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        if ($null -ne $originalLivePreview) { try { $excel.EnableLivePreview = $originalLivePreview } catch {} }
        if ($null -ne $originalMacroAnimations) { try { $excel.EnableMacroAnimations = $originalMacroAnimations } catch {} }
    }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    if (Test-Path -LiteralPath $probeRoot) { Remove-Item -LiteralPath $probeRoot -Recurse -Force }
}
