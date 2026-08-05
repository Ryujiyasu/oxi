# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-deftype-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$excel = $null
$workbook = $null
$sheet = $null
$component = $null

$source = @'
Option Explicit
DefInt A-C
DefDbl D, X-Y
DefStr S
DefObj O
DefVar V-W, Z

Public Sub OxiDefTypeProbe()
    Dim apple, cat, dog, xray, sample, objectValue, variantValue, zebra
    With ThisWorkbook.Worksheets(1)
        .Cells(1, 1).Value2 = VarType(apple)
        .Cells(2, 1).Value2 = VarType(cat)
        .Cells(3, 1).Value2 = VarType(dog)
        .Cells(4, 1).Value2 = VarType(xray)
        .Cells(5, 1).Value2 = VarType(sample)
        .Cells(6, 1).Value2 = VarType(objectValue)
        .Cells(7, 1).Value2 = VarType(variantValue)
        .Cells(8, 1).Value2 = VarType(zebra)
    End With
End Sub
'@

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($workbookPath, 52)
    try {
        $component = $workbook.VBProject.VBComponents.Add(1)
    }
    catch {
        throw "Excel COM is available, but VBProject access is disabled. $($_.Exception.Message)"
    }
    if ($null -eq $component) {
        throw 'Excel returned no VBComponent. Enable trust access to the VBA project object model.'
    }
    $component.Name = 'OxiDefTypeModule'
    $component.CodeModule.AddFromString($source)
    $excel.Run("'$($workbook.Name)'!OxiDefTypeModule.OxiDefTypeProbe")
    $sheet = $workbook.Worksheets.Item(1)

    $result = [ordered]@{
        a_integer = [long]$sheet.Cells.Item(1, 1).Value2
        c_integer = [long]$sheet.Cells.Item(2, 1).Value2
        d_double = [long]$sheet.Cells.Item(3, 1).Value2
        x_double = [long]$sheet.Cells.Item(4, 1).Value2
        s_string = [long]$sheet.Cells.Item(5, 1).Value2
        o_object = [long]$sheet.Cells.Item(6, 1).Value2
        v_variant = [long]$sheet.Cells.Item(7, 1).Value2
        z_variant = [long]$sheet.Cells.Item(8, 1).Value2
    }
    $expected = [ordered]@{
        a_integer = 2L
        c_integer = 2L
        d_double = 5L
        x_double = 5L
        s_string = 8L
        o_object = 9L
        v_variant = 0L
        z_variant = 0L
    }
    foreach ($key in $expected.Keys) {
        if ($result[$key] -ne $expected[$key]) {
            throw "COM result mismatch for ${key}: expected '$($expected[$key])', got '$($result[$key])'"
        }
    }
    $result | ConvertTo-Json
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($comObject in @($sheet, $component, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    if (Test-Path -LiteralPath $probeRoot) {
        [IO.Directory]::Delete($probeRoot, $true)
    }
}
