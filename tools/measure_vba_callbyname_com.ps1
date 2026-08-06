# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-callbyname-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'callbyname-probe.xlsm'
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

    $classComponent = $workbook.VBProject.VBComponents.Add(2)
    $classComponent.Name = 'CallByNameProbeClass'
    $classComponent.CodeModule.AddFromString(@'
Option Explicit

Private storedValue As Long

Public Property Let Value(ByVal newValue As Long)
    storedValue = newValue
End Property

Public Property Get Value() As Long
    Value = storedValue
End Property
'@)

    $standardComponent = $workbook.VBProject.VBComponents.Add(1)
    $standardComponent.Name = 'CallByNameProbe'
    $standardComponent.CodeModule.AddFromString(@'
Option Explicit

Public Function ExerciseCallByName() As Long
    Dim target As New CallByNameProbeClass
    CallByName target, "Value", VbLet, 40
    ExerciseCallByName = VBA.CallByName(target, "Value", VbGet) + 2
End Function
'@)

    $actual = [long]$excel.Run("'$($workbook.Name)'!CallByNameProbe.ExerciseCallByName")
    if ($actual -ne 42) {
        throw "CallByName COM execution mismatch: expected 42, got $actual"
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
    foreach ($expected in @(
        '[CallByNameProbeClass]',
        '[CallByNameProbe]',
        'procedures: 1, statements: 3, max nesting: 0, unparsed: 0',
        'CallByName: dispatches an object member by name; target resolution requires runtime type context',
        'VBA.CallByName: dispatches an object member by name; target resolution requires runtime type context'
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'CallByNameProbe' })[0]
    $dynamicFindings = @($probeModule.findings | Where-Object { $_.reason -like '*runtime type context*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.CallByName -ne 1 -or
        $probeModule.api_names.'VBA.CallByName' -ne 1 -or
        $dynamicFindings.Count -ne 2) {
        throw 'CallByName JSON analysis did not preserve dynamic member dispatch diagnostics'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'CallByName JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM CallByName returned: $actual"
    $analysis.TrimEnd()
    Write-Output 'VBA CallByName COM execution and analysis: PASS'
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
