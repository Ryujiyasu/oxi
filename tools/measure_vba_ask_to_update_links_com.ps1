# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-ask-update-links-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'ask-update-links-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalAskToUpdateLinks = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $originalAskToUpdateLinks = [bool]$excel.AskToUpdateLinks
    $workbook = $excel.Workbooks.Add()

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'AskToUpdateLinksProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableLinkPrompts()
    Application.AskToUpdateLinks = False
End Sub

Public Sub EnableLinkPrompts()
    Application.AskToUpdateLinks = True
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!AskToUpdateLinksProbe.DisableLinkPrompts")
    $disabled = [bool]$excel.AskToUpdateLinks
    if ($disabled) {
        throw 'AskToUpdateLinks COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!AskToUpdateLinksProbe.EnableLinkPrompts")
    $enabled = [bool]$excel.AskToUpdateLinks
    if (-not $enabled) {
        throw 'AskToUpdateLinks COM enabled-state mismatch: expected True'
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.AskToUpdateLinks = $originalAskToUpdateLinks
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
        '[AskToUpdateLinksProbe]',
        'procedures: 2, statements: 2, max nesting: 0, unparsed: 0',
        "Application.AskToUpdateLinks: changes Excel's process-global prompt policy for updating external links"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'AskToUpdateLinksProbe' })[0]
    $promptFindings = @($probeModule.findings | Where-Object { $_.reason -like '*prompt policy for updating external links*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.AskToUpdateLinks' -ne 2 -or
        $promptFindings.Count -ne 2) {
        throw 'AskToUpdateLinks JSON analysis did not preserve global link-prompt dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'AskToUpdateLinks JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM AskToUpdateLinks states: original=$originalAskToUpdateLinks disabled=$disabled enabled=$enabled"
    $analysis.TrimEnd()
    Write-Output 'VBA AskToUpdateLinks COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel -and $null -ne $originalAskToUpdateLinks) {
        try { $excel.AskToUpdateLinks = $originalAskToUpdateLinks } catch {}
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
