# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-enable-events-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'enable-events-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)

    $sheetComponent = $workbook.VBProject.VBComponents.Item($sheet.CodeName)
    $sheetComponent.CodeModule.AddFromString(@'
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 Then
        Me.Range("B1").Value2 = Me.Range("B1").Value2 + 1
    End If
End Sub
'@)

    $standardComponent = $workbook.VBProject.VBComponents.Add(1)
    $standardComponent.Name = 'EnableEventsProbe'
    $standardComponent.CodeModule.AddFromString(@'
Option Explicit

Public Sub ExerciseEnableEvents()
    Application.EnableEvents = False
    Sheet1.Range("A1").Value2 = 1
    Application.EnableEvents = True
    Sheet1.Range("A2").Value2 = 1
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!EnableEventsProbe.ExerciseEnableEvents")
    $actual = [long]$sheet.Range('B1').Value2
    if ($actual -ne 1) {
        throw "EnableEvents COM execution mismatch: expected one Worksheet_Change event, got $actual"
    }

    $workbook.SaveAs($workbookPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.EnableEvents = $true
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
        '[EnableEventsProbe]',
        'procedures: 1, statements: 4, max nesting: 0, unparsed: 0',
        "Application.EnableEvents: changes Excel's process-global event delivery state"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'EnableEventsProbe' })[0]
    $eventFindings = @($probeModule.findings | Where-Object { $_.reason -like '*event delivery state*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.EnableEvents' -ne 2 -or
        $eventFindings.Count -ne 2) {
        throw 'EnableEvents JSON analysis did not preserve global event-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'EnableEvents JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM Worksheet_Change count: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA EnableEvents COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        try { $excel.EnableEvents = $true } catch {}
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
