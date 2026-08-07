# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-random-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'random-probe.xlsm'
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
    $component.Name = 'RandomProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Function SeededSequenceRepeats() As Boolean
    Dim first As Single, second As Single, resetValue As Single
    resetValue = Rnd(-1)
    Randomize 42
    first = Rnd
    resetValue = Rnd(-1)
    Randomize 42
    second = Rnd
    SeededSequenceRepeats = (first = second)
End Function
'@)

    $actual = [bool]$excel.Run("'$($workbook.Name)'!RandomProbe.SeededSequenceRepeats")
    if (-not $actual) {
        throw 'Rnd/Randomize COM execution mismatch: fixed seed did not reproduce the sequence'
    }

    $workbook.SaveAs($workbookPath, 52)
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
        '[RandomProbe]',
        'procedures: 1, statements: 8, max nesting: 0, unparsed: 0',
        "Rnd: uses VBA's process-global pseudorandom generator; results depend on seed and call order",
        "Randomize: uses VBA's process-global pseudorandom generator; results depend on seed and call order"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'RandomProbe' })[0]
    $randomFindings = @($probeModule.findings | Where-Object { $_.reason -like '*seed and call order*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.Rnd -ne 4 -or
        $probeModule.api_names.Randomize -ne 2 -or
        $randomFindings.Count -ne 6) {
        throw 'Rnd/Randomize JSON analysis did not preserve generator-state dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Rnd/Randomize JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM fixed-seed sequence repeated: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA Rnd/Randomize COM execution and analysis: PASS'
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
