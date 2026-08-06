# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-application-run-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'application-run-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'RunProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub InvokeRun()
    Call Application.Run("'application-run-probe.xlsm'!RunProbe.RunTarget", 40)
End Sub

Public Sub RunTarget(ByVal number As Long)
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = number + 2
End Sub

Private Function HiddenUnused() As Long
    HiddenUnused = 99
End Function
'@)

    $workbook.SaveAs($workbookPath, 52)
    $excel.Run("'$($workbook.Name)'!RunProbe.InvokeRun")
    $actual = [long]$workbook.Worksheets.Item(1).Range('A1').Value2
    if ($actual -ne 42) {
        throw "Application.Run COM execution mismatch: expected 42, got $actual"
    }

    $workbook.Save()
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
    foreach ($expected in @(
        '[RunProbe]',
        'procedures: 3, statements: 3, max nesting: 0, unparsed: 0',
        'Application.Run: dispatches a macro by name; target resolution requires workbook context'
    )) {
        if (-not $analysis.Contains($expected)) {
            throw "Expected analysis fragment not found: $expected`n$analysis"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'RunProbe' })[0]
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        @($probeModule.uncalled_procedures).Count -ne 1 -or
        $probeModule.uncalled_procedures[0] -ne 'HiddenUnused' -or
        $probeModule.api_names.'Application.Run' -ne 1) {
        throw 'Application.Run JSON analysis did not preserve dynamic dispatch diagnostics'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.Run JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Application.Run returned: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.Run COM execution and conservative analysis: PASS'
}
finally {
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
