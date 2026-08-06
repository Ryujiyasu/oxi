# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-xlsm-analysis-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$workbookCopyPath = Join-Path $probeRoot 'probe-copy.xlsm'
$workbookVariantPath = Join-Path $probeRoot 'probe-variant.xlsm'
$repoRoot = Split-Path -Parent $PSScriptRoot
$cliPath = Join-Path $repoRoot 'target\debug\oxidocs.exe'
$excel = $null
$workbook = $null
$module = $null
$sheet = $null

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'AnalysisProbe'
    $module.CodeModule.AddFromString(@'
Option Explicit
#If Win64 Then
Private Const PlatformBits As Long = 64
#Else
Private Const PlatformBits As Long = 32
#End If

Public Sub BuildReport(ByVal target As Worksheet)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = SuffixValue$(42&)
    ws.Range("B1").Value = Len(SuffixValue$(42&))
    ws.Range("A1").Font.Bold = True
    ws.Range("C1").Value = WideValue()
    ws.Range("D1").Value = PlatformBits
    ws.Range("E1").Value = ParenthesesValue()
End Sub

Private Function HiddenHelper(ByVal value As Long) As Long
    HiddenHelper = value + 1
End Function

Private Function SuffixValue$(ByVal number&)
    Dim padded As String * 4
    Dim chunks() As String
    Dim lazy As New Collection
    ReDim chunks(1 To 2) As String
    chunks(1) = CStr(number&)
    lazy.Add chunks(1)
    padded = lazy(1)
    SuffixValue$ = padded
End Function

Private Function WideValue^()
    WideValue = 42^ + 2
End Function

Private Function ParenthesesValue$()
    Dim direct As Long
    Dim wrapped As Long
    Dim explicitGrouped As Long
    direct = 1
    wrapped = 1
    explicitGrouped = 1
    Mutate direct
    Mutate (wrapped)
    Call Mutate((explicitGrouped))
    LabelBranch
    ParenthesesValue = "D" & CStr(direct) & "W" & CStr(wrapped) & "G" & CStr(explicitGrouped)
End Function

Private Sub Mutate(ByRef value As Long)
    value = 99
End Sub

Private Sub LabelBranch()
    GoTo LabelCollision
LabelCollision:
End Sub

Private Sub LabelCollision()
End Sub

Private Function SelfNamedValue() As Long
    SelfNamedValue = 7
End Function

Private Sub DeadChainA()
    DeadChainB
End Sub

Private Sub DeadChainB()
End Sub
'@)
    $sheet = $workbook.Worksheets.Item(1)
    $excel.Run("'$($workbook.Name)'!AnalysisProbe.BuildReport", $sheet)
    $actualValue = [string]$sheet.Range('A1').Value2
    $fixedLength = [long]$sheet.Range('B1').Value2
    $wideValue = [long]$sheet.Range('C1').Value2
    $platformBits = [long]$sheet.Range('D1').Value2
    $parenthesesValue = [string]$sheet.Range('E1').Value2
    if ($actualValue -ne '42' -or $fixedLength -ne 4 -or $wideValue -ne 44 -or $platformBits -ne 64 -or $parenthesesValue -ne 'D99W1G1' -or -not [bool]$sheet.Range('A1').Font.Bold) {
        throw "COM execution mismatch: value=[$actualValue], fixed-length=$fixedLength, wide=$wideValue, platform=$platformBits, parentheses=[$parenthesesValue], bold=$([bool]$sheet.Range('A1').Font.Bold)"
    }
    $workbook.SaveAs($workbookPath, 52)
    $workbook.SaveCopyAs($workbookCopyPath)
    $module.CodeModule.ReplaceLine(20, '    HiddenHelper = value + 2')
    $module.CodeModule.InsertLines(2, 'Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr')
    $module.CodeModule.AddFromString(@'

Private Function VariantOnly(ByVal value As Long) As Long
    VariantOnly = SetTimer(0, 0, value, AddressOf HiddenHelper)
End Function
'@)
    $workbook.SaveAs($workbookVariantPath, 52)
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    & cargo build -q -p oxidocs-cli --manifest-path (Join-Path $repoRoot 'Cargo.toml')
    if ($LASTEXITCODE -ne 0) {
        throw "oxidocs-cli build failed with exit code $LASTEXITCODE"
    }

    $output = (& $cliPath vba-analyze $workbookPath | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-analyze failed with exit code $LASTEXITCODE"
    }

    $expectations = @(
        '[AnalysisProbe]',
        'verdict: A (report generation)',
        'procedures: 11, statements: 34',
        'unparsed: 0',
        '#If Win64 Then: conditional compilation; the source differs by build',
        '#End If: conditional compilation; the source differs by build',
        'Summary:'
    )
    foreach ($expected in $expectations) {
        if (-not $output.Contains($expected)) {
            throw "Expected output fragment not found: $expected`n$output"
        }
    }

    $output.TrimEnd()
    $inventory = (& $cliPath vba-inventory $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory failed with exit code $LASTEXITCODE"
    }
    $inventoryExpectations = @(
        'Structurally identical modules (standard fingerprint):',
        'probe.xlsm::AnalysisProbe',
        'probe-copy.xlsm::AnalysisProbe',
        'Related modules (standard fingerprint):',
        'probe-variant.xlsm::AnalysisProbe',
        '76.9% (shared 10; only 1/2; diverged: HiddenHelper; declarations differ)',
        'Inventory: 3 succeeded, 0 failed'
    )
    foreach ($expected in $inventoryExpectations) {
        if (-not $inventory.Contains($expected)) {
            throw "Expected inventory fragment not found: $expected`n$inventory"
        }
    }
    $inventory.TrimEnd()

    $jsonOutput = (& $cliPath vba-inventory-json $probeRoot | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "vba-inventory-json failed with exit code $LASTEXITCODE"
    }
    $jsonReport = $jsonOutput | ConvertFrom-Json
    if ($jsonReport.schema -ne 'oxivba-inventory-v1') {
        throw "Unexpected JSON schema: $($jsonReport.schema)"
    }
    if (@($jsonReport.projects).Count -ne 3) {
        throw "Expected 3 JSON projects, got $(@($jsonReport.projects).Count)"
    }
    if (@($jsonReport.duplicate_groups).Count -ne 1) {
        throw "Expected 1 JSON duplicate group, got $(@($jsonReport.duplicate_groups).Count)"
    }
    if (@($jsonReport.related_modules).Count -ne 2) {
        throw "Expected 2 JSON related pairs, got $(@($jsonReport.related_modules).Count)"
    }
    $divergedNames = @($jsonReport.related_modules | ForEach-Object { $_.diverged })
    if ($divergedNames -notcontains 'HiddenHelper') {
        throw "JSON related pairs did not report HiddenHelper divergence"
    }
    if (@($jsonReport.related_modules | Where-Object { $_.declarations_differ }).Count -ne 2) {
        throw "JSON related pairs did not report the changed Declare context"
    }
    $variantProject = @($jsonReport.projects | Where-Object { $_.path -like '*probe-variant.xlsm' })[0]
    $variantModule = @($variantProject.modules | Where-Object { $_.name -eq 'AnalysisProbe' })[0]
    if ($variantModule.metrics.unparsed -ne 0) {
        throw "AddressOf variant produced unparsed source"
    }
    if (@($variantModule.external_declares) -notcontains 'SetTimer') {
        throw "AddressOf variant did not report its SetTimer declaration"
    }
    if ($variantModule.api_names.Worksheet -ne 2) {
        throw "Procedure and local declaration types were not both reported"
    }
    if ($variantModule.api_names.Collection -ne 1) {
        throw "As New declaration type was not reported"
    }
    if (@($variantModule.uncalled_procedures) -notcontains 'LabelCollision') {
        throw "GoTo label incorrectly counted as a call to LabelCollision"
    }
    if (@($variantModule.uncalled_procedures) -notcontains 'SelfNamedValue') {
        throw "Function result assignment incorrectly counted as a call to SelfNamedValue"
    }
    if (@($variantModule.uncalled_procedures) -notcontains 'DeadChainA' -or @($variantModule.uncalled_procedures) -notcontains 'DeadChainB') {
        throw "Dead private call chain was not fully reported as uncalled"
    }
    if (@($variantModule.uncalled_procedures) -contains 'LabelBranch') {
        throw "Called LabelBranch was incorrectly reported as uncalled"
    }
    if (@($jsonReport.errors).Count -ne 0) {
        throw "JSON inventory reported unexpected errors"
    }
    Write-Output 'JSON inventory schema round-trip: PASS'
    Write-Output 'COM XLSM extraction and analysis probe: PASS'
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($comObject in @($sheet, $module, $workbook, $excel)) {
        if ($null -ne $comObject) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $probeRoot) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}
