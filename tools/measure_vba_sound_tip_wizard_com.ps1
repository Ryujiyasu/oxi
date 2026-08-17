# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-sound-tip-wizard-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'sound-tip-wizard-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalSound = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalSound = [bool]$excel.EnableSound
    $originalTipWizard = [bool]$excel.EnableTipWizard

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'SoundTipWizardProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableSound()
    Application.EnableSound = False
End Sub

Public Sub EnableSound()
    Application.EnableSound = True
End Sub

Public Function DisableTipWizard() As Long
    On Error Resume Next
    Application.EnableTipWizard = False
    DisableTipWizard = Err.Number
End Function

Public Function EnableTipWizard() As Long
    On Error Resume Next
    Application.EnableTipWizard = True
    EnableTipWizard = Err.Number
End Function
'@)

    $excel.Run("'$($workbook.Name)'!SoundTipWizardProbe.DisableSound")
    $soundDisabled = [bool]$excel.EnableSound
    $excel.Run("'$($workbook.Name)'!SoundTipWizardProbe.EnableSound")
    $soundEnabled = [bool]$excel.EnableSound
    if ($soundDisabled -or -not $soundEnabled) {
        throw 'Application.EnableSound COM state did not follow False/True assignments'
    }

    $tipDisableError = [int]$excel.Run("'$($workbook.Name)'!SoundTipWizardProbe.DisableTipWizard")
    $tipEnableError = [int]$excel.Run("'$($workbook.Name)'!SoundTipWizardProbe.EnableTipWizard")
    $tipAfter = [bool]$excel.EnableTipWizard
    if ($tipDisableError -eq 0 -or $tipEnableError -eq 0 -or $tipAfter) {
        throw 'Application.EnableTipWizard did not preserve the measured disabled/rejected behavior'
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.EnableSound = $originalSound
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    & cargo build -q -p oxidocs-cli --manifest-path (Join-Path $repoRoot 'Cargo.toml')
    if ($LASTEXITCODE -ne 0) { throw "oxidocs-cli build failed with exit code $LASTEXITCODE" }
    $analysis = (& $cliPath vba-analyze $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "vba-analyze failed with exit code $LASTEXITCODE" }
    foreach ($expectedFragment in @(
        '[SoundTipWizardProbe]',
        'procedures: 4, statements: 8, max nesting: 0, unparsed: 0',
        'verdict: D',
        "Application.EnableSound: changes Excel's process-global user-interface sound state",
        "Application.EnableTipWizard: requests Excel's legacy TipWizard user interface; modern Excel reports it disabled and rejects assignments with error 1004"
    )) {
        if (-not $analysis.Contains($expectedFragment)) { throw "Expected analysis fragment not found: $expectedFragment`n$analysis" }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "vba-inventory-json failed with exit code $LASTEXITCODE" }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'SoundTipWizardProbe' })[0]
    $soundFindings = @($probeModule.findings | Where-Object { $_.reason -like '*user-interface sound state*' })
    $tipFindings = @($probeModule.findings | Where-Object { $_.reason -like '*legacy TipWizard*' })
    if ($null -eq $probeModule -or $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.EnableSound' -ne 2 -or
        $probeModule.api_names.'Application.EnableTipWizard' -ne 2 -or
        $soundFindings.Count -ne 2 -or $tipFindings.Count -ne 2) {
        throw 'Sound/TipWizard JSON analysis did not preserve both UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) { throw 'Sound/TipWizard JSON inventory reported unexpected errors' }

    Write-Output "Excel COM sound/TipWizard: sound original=$originalSound false=$soundDisabled true=$soundEnabled; TipWizard original=$originalTipWizard disableError=$tipDisableError enableError=$tipEnableError after=$tipAfter"
    $analysis.TrimEnd()
    Write-Output 'VBA sound and TipWizard COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalSound) { try { $excel.EnableSound = $originalSound } catch {} }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    if (Test-Path -LiteralPath $probeRoot) { Remove-Item -LiteralPath $probeRoot -Recurse -Force }
}
