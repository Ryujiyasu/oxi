# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-property-" + [guid]::NewGuid().ToString('N'))
$basePath = Join-Path $probeRoot 'property-base.xlsm'
$variantPath = Join-Path $probeRoot 'property-variant.xlsm'
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

    $propertyModule = $workbook.VBProject.VBComponents.Add(2)
    $propertyModule.Name = 'PropertyProbe'
    $propertyModule.CodeModule.AddFromString(@'
Option Explicit
Private mValue As Long

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As Long)
    mValue = newValue
End Property
'@)

    $entryModule = $workbook.VBProject.VBComponents.Add(1)
    $entryModule.Name = 'PropertyEntry'
    $entryModule.CodeModule.AddFromString(@'
Option Explicit

Public Sub RunProbe()
    Dim item As New PropertyProbe
    item.Value = 7
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = item.Value
End Sub
'@)

    $excel.Run("'$($workbook.Name)'!PropertyEntry.RunProbe")
    $actual = [long]$workbook.Worksheets.Item(1).Range('A1').Value2
    if ($actual -ne 7) {
        throw "Property COM execution mismatch: expected 7, got $actual"
    }

    $workbook.SaveAs($basePath, 52)
    $propertyModule.CodeModule.ReplaceLine(9, '    mValue = newValue + 1')
    $workbook.SaveAs($variantPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    & cargo build -q -p oxidocs-cli --manifest-path (Join-Path $repoRoot 'Cargo.toml')
    if ($LASTEXITCODE -ne 0) {
        throw "oxidocs-cli build failed with exit code $LASTEXITCODE"
    }

    $inventory = (& $cliPath vba-inventory $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory failed with exit code $LASTEXITCODE"
    }
    foreach ($expected in @(
        'property-base.xlsm::PropertyProbe',
        'property-variant.xlsm::PropertyProbe',
        '33.3% (shared 1; only 1/1; diverged: Value (Property Let))',
        'Inventory: 2 succeeded, 0 failed'
    )) {
        if (-not $inventory.Contains($expected)) {
            throw "Expected inventory fragment not found: $expected`n$inventory"
        }
    }

    $jsonOutput = (& $cliPath vba-inventory-json $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    if (@($jsonReport.projects).Count -ne 2 -or @($jsonReport.errors).Count -ne 0) {
        throw 'Property JSON inventory did not analyze both workbooks cleanly'
    }
    $propertyPair = @($jsonReport.related_modules | Where-Object {
        $_.left -like '*::PropertyProbe' -and $_.right -like '*::PropertyProbe'
    })[0]
    if ($null -eq $propertyPair -or @($propertyPair.diverged) -notcontains 'Value (Property Let)') {
        throw 'Property JSON comparison did not isolate the changed Let accessor'
    }
    if ($propertyPair.shared -ne 1 -or $propertyPair.only_left -ne 1 -or $propertyPair.only_right -ne 1) {
        throw 'Property JSON comparison returned incorrect accessor overlap'
    }

    $inventory.TrimEnd()
    Write-Output 'Property Get/Let COM execution and fingerprint comparison: PASS'
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
