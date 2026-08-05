# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-options-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$excel = $null
$workbook = $null
$sheet = $null
$baseZero = $null
$baseOne = $null
$binary = $null
$textCompare = $null
$runner = $null

$baseZeroSource = @'
Option Explicit
Option Base 0
Public Function BaseZeroDeclared() As Long
    Dim values(3) As Long
    BaseZeroDeclared = LBound(values)
End Function
Public Function BaseZeroArrayFunction() As Long
    BaseZeroArrayFunction = LBound(Array("a", "b"))
End Function
'@

$baseOneSource = @'
Option Explicit
Option Base 1
Public Function BaseOneDeclared() As Long
    Dim values(3) As Long
    BaseOneDeclared = LBound(values)
End Function
Public Function BaseOneArrayFunction() As Long
    BaseOneArrayFunction = LBound(Array("a", "b"))
End Function
'@

$binarySource = @'
Option Explicit
Option Compare Binary
Public Function BinaryEqual() As Boolean
    BinaryEqual = ("a" = "A")
End Function
Public Function BinaryLike() As Boolean
    BinaryLike = ("a" Like "[A-Z]")
End Function
'@

$textSource = @'
Option Explicit
Option Compare Text
Public Function TextEqual() As Boolean
    TextEqual = ("a" = "A")
End Function
Public Function TextLike() As Boolean
    TextLike = ("a" Like "[A-Z]")
End Function
'@

$runnerSource = @'
Option Explicit
Public Sub OxiOptionsProbe()
    With ThisWorkbook.Worksheets(1)
        .Cells(1, 1).Value2 = BaseZeroDeclared()
        .Cells(2, 1).Value2 = BaseOneDeclared()
        .Cells(3, 1).Value2 = BaseZeroArrayFunction()
        .Cells(4, 1).Value2 = BaseOneArrayFunction()
        .Cells(5, 1).Value2 = BinaryEqual()
        .Cells(6, 1).Value2 = TextEqual()
        .Cells(7, 1).Value2 = BinaryLike()
        .Cells(8, 1).Value2 = TextLike()
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
        $baseZero = $workbook.VBProject.VBComponents.Add(1)
        $baseOne = $workbook.VBProject.VBComponents.Add(1)
        $binary = $workbook.VBProject.VBComponents.Add(1)
        $textCompare = $workbook.VBProject.VBComponents.Add(1)
        $runner = $workbook.VBProject.VBComponents.Add(1)
    }
    catch {
        throw "Excel COM is available, but VBProject access is disabled. $($_.Exception.Message)"
    }
    $baseZero.Name = 'OxiBaseZero'
    $baseOne.Name = 'OxiBaseOne'
    $binary.Name = 'OxiCompareBinary'
    $textCompare.Name = 'OxiCompareText'
    $runner.Name = 'OxiOptionsModule'
    $baseZero.CodeModule.AddFromString($baseZeroSource)
    $baseOne.CodeModule.AddFromString($baseOneSource)
    $binary.CodeModule.AddFromString($binarySource)
    $textCompare.CodeModule.AddFromString($textSource)
    $runner.CodeModule.AddFromString($runnerSource)
    $excel.Run("'$($workbook.Name)'!OxiOptionsModule.OxiOptionsProbe")
    $sheet = $workbook.Worksheets.Item(1)

    $result = [ordered]@{
        base_zero_declared = [long]$sheet.Cells.Item(1, 1).Value2
        base_one_declared = [long]$sheet.Cells.Item(2, 1).Value2
        base_zero_array_function = [long]$sheet.Cells.Item(3, 1).Value2
        base_one_array_function = [long]$sheet.Cells.Item(4, 1).Value2
        binary_equal = [bool]$sheet.Cells.Item(5, 1).Value2
        text_equal = [bool]$sheet.Cells.Item(6, 1).Value2
        binary_like = [bool]$sheet.Cells.Item(7, 1).Value2
        text_like = [bool]$sheet.Cells.Item(8, 1).Value2
    }
    $expected = [ordered]@{
        base_zero_declared = 0L
        base_one_declared = 1L
        base_zero_array_function = 0L
        base_one_array_function = 1L
        binary_equal = $false
        text_equal = $true
        binary_like = $false
        text_like = $true
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
    foreach ($comObject in @($sheet, $runner, $textCompare, $binary, $baseOne, $baseZero, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    if (Test-Path -LiteralPath $probeRoot) {
        [IO.Directory]::Delete($probeRoot, $true)
    }
}
