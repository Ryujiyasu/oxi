# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-move-after-return-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'move-after-return-probe.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$originalMoveAfterReturn = $null
$originalDirection = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $originalMoveAfterReturn = [bool]$excel.MoveAfterReturn
    $originalDirection = [int]$excel.MoveAfterReturnDirection

    $component = $workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'MoveAfterReturnProbe'
    $component.CodeModule.AddFromString(@'
Option Explicit

Public Sub DisableMoveAfterReturn()
    Application.MoveAfterReturn = False
End Sub

Public Sub EnableMoveAfterReturn()
    Application.MoveAfterReturn = True
End Sub

Public Sub MoveDownAfterReturn()
    Application.MoveAfterReturnDirection = xlDown
End Sub

Public Sub MoveRightAfterReturn()
    Application.MoveAfterReturnDirection = xlToRight
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!MoveAfterReturnProbe.DisableMoveAfterReturn")
    $disabled = [bool]$excel.MoveAfterReturn
    if ($disabled) {
        throw 'Application.MoveAfterReturn COM disabled-state mismatch: expected False'
    }

    $excel.Run("'$($workbook.Name)'!MoveAfterReturnProbe.EnableMoveAfterReturn")
    $enabled = [bool]$excel.MoveAfterReturn
    if (-not $enabled) {
        throw 'Application.MoveAfterReturn COM enabled-state mismatch: expected True'
    }

    $excel.Run("'$($workbook.Name)'!MoveAfterReturnProbe.MoveDownAfterReturn")
    $down = [int]$excel.MoveAfterReturnDirection
    if ($down -ne -4121) {
        throw "Application.MoveAfterReturnDirection COM xlDown mismatch: expected -4121, got $down"
    }

    $excel.Run("'$($workbook.Name)'!MoveAfterReturnProbe.MoveRightAfterReturn")
    $right = [int]$excel.MoveAfterReturnDirection
    if ($right -ne -4161) {
        throw "Application.MoveAfterReturnDirection COM xlToRight mismatch: expected -4161, got $right"
    }

    $workbook.SaveAs($workbookPath, 52)
    $excel.MoveAfterReturnDirection = $originalDirection
    $excel.MoveAfterReturn = $originalMoveAfterReturn
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
        '[MoveAfterReturnProbe]',
        'procedures: 4, statements: 4, max nesting: 0, unparsed: 0',
        'verdict: D',
        'Application.MoveAfterReturn: changes whether Excel moves the active selection after Enter',
        "Application.MoveAfterReturnDirection: changes Excel's process-global selection direction after Enter"
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
    $probeModule = @($jsonReport.projects[0].modules | Where-Object { $_.name -eq 'MoveAfterReturnProbe' })[0]
    $moveFindings = @($probeModule.findings | Where-Object { $_.reason -like '*active selection after Enter*' })
    $directionFindings = @($probeModule.findings | Where-Object { $_.reason -like '*selection direction after Enter*' })
    if ($null -eq $probeModule -or
        $probeModule.metrics.unparsed -ne 0 -or
        $probeModule.api_names.'Application.MoveAfterReturn' -ne 2 -or
        $probeModule.api_names.'Application.MoveAfterReturnDirection' -ne 2 -or
        $moveFindings.Count -ne 2 -or
        $directionFindings.Count -ne 2) {
        throw 'Application.MoveAfterReturn JSON analysis did not preserve global UI dependencies'
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw 'Application.MoveAfterReturn JSON inventory reported unexpected errors'
    }

    Write-Output "Excel COM MoveAfterReturn values: original=$originalMoveAfterReturn disabled=$disabled enabled=$enabled directionOriginal=$originalDirection down=$down right=$right"
    $analysis.TrimEnd()
    Write-Output 'VBA Application.MoveAfterReturn COM execution and analysis: PASS'
}
finally {
    if ($null -ne $excel) {
        if ($null -ne $originalDirection) {
            try { $excel.MoveAfterReturnDirection = $originalDirection } catch {}
        }
        if ($null -ne $originalMoveAfterReturn) {
            try { $excel.MoveAfterReturn = $originalMoveAfterReturn } catch {}
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
