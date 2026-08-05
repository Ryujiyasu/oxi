# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

param(
    [switch]$KeepArtifacts
)

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-file-io-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$excel = $null
$workbook = $null
$sheet = $null
$component = $null

$source = @'
Option Explicit

Public Sub OxiFileIoProbe()
    Dim printNo As Integer
    Dim writeNo As Integer
    Dim readNo As Integer
    Dim leftNo As Integer
    Dim rightNo As Integer
    Dim binaryNo As Integer
    Dim optionalNo As Integer
    Dim lockNo As Integer
    Dim widthNo As Integer
    Dim resetNo As Integer
    Dim first As String
    Dim second As Long
    Dim third As Boolean
    Dim wholeLine As String
    Dim printPath As String
    Dim writePath As String
    Dim leftPath As String
    Dim rightPath As String
    Dim binaryPath As String
    Dim optionalPath As String
    Dim copiedPath As String
    Dim renamedPath As String
    Dim directoryPath As String
    Dim originalDir As String
    Dim lockPath As String
    Dim widthPath As String
    Dim resetPath As String
    Dim sourceLong As Long
    Dim readLong As Long
    Dim optionalLong As Long

    printPath = ThisWorkbook.Path & "\print.txt"
    writePath = ThisWorkbook.Path & "\write.txt"
    leftPath = ThisWorkbook.Path & "\left.txt"
    rightPath = ThisWorkbook.Path & "\right.txt"
    binaryPath = ThisWorkbook.Path & "\binary.dat"
    optionalPath = ThisWorkbook.Path & "\optional-hash.dat"
    copiedPath = ThisWorkbook.Path & "\copied.txt"
    renamedPath = ThisWorkbook.Path & "\renamed.txt"
    directoryPath = ThisWorkbook.Path & "\created-dir"
    lockPath = ThisWorkbook.Path & "\lock.dat"
    widthPath = ThisWorkbook.Path & "\width.txt"
    resetPath = ThisWorkbook.Path & "\reset.txt"

    printNo = FreeFile
    Open printPath For Output Access Write Lock Write As #printNo
    Print #printNo, "alpha"; 42;
    Print #printNo, "omega"
    Close #printNo

    writeNo = FreeFile
    Open writePath For Output Access Write Lock Write As #writeNo
    Write #writeNo, "alpha", 42, True
    Close #writeNo

    readNo = FreeFile
    Open writePath For Input Access Read Lock Read As #readNo
    Input #readNo, first, second, third
    Close #readNo

    readNo = FreeFile
    Open printPath For Input Access Read Lock Read As #readNo
    Line Input #readNo, wholeLine
    Close #readNo

    leftNo = FreeFile
    Open leftPath For Output As #leftNo
    rightNo = FreeFile
    Open rightPath For Output As #rightNo
    Close #leftNo, #rightNo

    binaryNo = FreeFile
    Open binaryPath For Binary Access Read Write As #binaryNo
    sourceLong = &H12345678
    Put #binaryNo, 1, sourceLong

    With ThisWorkbook.Worksheets(1)
        .Cells(1, 1).Value2 = wholeLine
        .Cells(2, 1).Value2 = first
        .Cells(3, 1).Value2 = second
        .Cells(4, 1).Value2 = third
        .Cells(5, 1).Value2 = "multi-close-ok"
        .Cells(6, 1).Value2 = LOF(binaryNo)
        .Cells(7, 1).Value2 = Loc(binaryNo)
        .Cells(8, 1).Value2 = Seek(binaryNo)
        Seek #binaryNo, 1
        .Cells(9, 1).Value2 = Seek(binaryNo)
        Get #binaryNo, , readLong
        .Cells(10, 1).Value2 = readLong
        .Cells(11, 1).Value2 = EOF(binaryNo)
        .Cells(12, 1).Value2 = FileLen(binaryPath)
        .Cells(13, 1).Value2 = FreeFile
    End With
    Close #binaryNo
    ThisWorkbook.Worksheets(1).Cells(14, 1).Value2 = FileLen(binaryPath)

    optionalNo = FreeFile
    Open optionalPath For Binary As optionalNo
    Put optionalNo, 1, sourceLong
    Seek optionalNo, 1
    Get optionalNo, , optionalLong
    Close optionalNo
    ThisWorkbook.Worksheets(1).Cells(15, 1).Value2 = optionalLong

    FileCopy printPath, copiedPath
    ThisWorkbook.Worksheets(1).Cells(16, 1).Value2 = FileLen(copiedPath)
    Name copiedPath As renamedPath
    ThisWorkbook.Worksheets(1).Cells(17, 1).Value2 = (Dir$(copiedPath) = "")
    ThisWorkbook.Worksheets(1).Cells(18, 1).Value2 = (FileLen(renamedPath) > 0)
    SetAttr renamedPath, vbHidden
    ThisWorkbook.Worksheets(1).Cells(19, 1).Value2 = (GetAttr(renamedPath) And vbHidden) <> 0
    SetAttr renamedPath, vbNormal
    Kill renamedPath
    ThisWorkbook.Worksheets(1).Cells(20, 1).Value2 = (Dir$(renamedPath) = "")

    originalDir = CurDir$
    MkDir directoryPath
    ChDrive Left$(ThisWorkbook.Path, 1)
    ChDir directoryPath
    ThisWorkbook.Worksheets(1).Cells(21, 1).Value2 = CurDir$
    ChDrive Left$(originalDir, 1)
    ChDir originalDir
    RmDir directoryPath
    ThisWorkbook.Worksheets(1).Cells(22, 1).Value2 = (Dir$(directoryPath, vbDirectory) = "")

    lockNo = FreeFile
    Open lockPath For Binary As #lockNo
    Put #lockNo, 1, sourceLong
    Lock lockNo, 1 To 4
    Unlock #lockNo, 1 To 4
    Lock #lockNo, 2
    Unlock #lockNo, 2
    Lock #lockNo, To 4
    Unlock #lockNo, To 4
    Close #lockNo
    ThisWorkbook.Worksheets(1).Cells(23, 1).Value2 = "lock-unlock-ok"

    widthNo = FreeFile
    Open widthPath For Output As #widthNo
    Width #widthNo, 12
    Print #widthNo, "123456789012345"
    Close #widthNo
    ThisWorkbook.Worksheets(1).Cells(24, 1).Value2 = FileLen(widthPath)

    resetNo = FreeFile
    Open resetPath For Output As #resetNo
    Reset
    On Error Resume Next
    Open resetPath For Append As #resetNo
    ThisWorkbook.Worksheets(1).Cells(25, 1).Value2 = Err.Number
    If Err.Number = 0 Then Close #resetNo
    Err.Clear
    On Error GoTo 0
End Sub
'@

try {
    New-Item -ItemType Directory -Path $probeRoot | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # This affects only this isolated COM instance. The workbook and its VBA
    # source are both created below and removed when the probe completes.
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($workbookPath, 52)

    try {
        $component = $workbook.VBProject.VBComponents.Add(1)
    }
    catch {
        throw "Excel COM is available, but VBProject access is disabled. Enable 'Trust access to the VBA project object model', restart Excel, and rerun this probe. $($_.Exception.Message)"
    }
    if ($null -eq $component) {
        throw "Excel returned no VBComponent. Enable 'Trust access to the VBA project object model' and restart Excel."
    }

    $component.Name = 'OxiProbeModule'
    $component.CodeModule.AddFromString($source)
    $excel.Run("'$($workbook.Name)'!OxiProbeModule.OxiFileIoProbe")
    $workbook.Save()

    $sheet = $workbook.Worksheets.Item(1)
    $result = [ordered]@{
        excel_version = [string]$excel.Version
        line_input = [string]$sheet.Cells.Item(1, 1).Value2
        input_string = [string]$sheet.Cells.Item(2, 1).Value2
        input_long = [long]$sheet.Cells.Item(3, 1).Value2
        input_boolean = [bool]$sheet.Cells.Item(4, 1).Value2
        multiple_close = [string]$sheet.Cells.Item(5, 1).Value2
        binary_lof = [long]$sheet.Cells.Item(6, 1).Value2
        binary_loc_after_put = [long]$sheet.Cells.Item(7, 1).Value2
        binary_seek_after_put = [long]$sheet.Cells.Item(8, 1).Value2
        binary_seek_after_reset = [long]$sheet.Cells.Item(9, 1).Value2
        binary_get_long = [long]$sheet.Cells.Item(10, 1).Value2
        binary_eof_after_get = [bool]$sheet.Cells.Item(11, 1).Value2
        file_len = [long]$sheet.Cells.Item(12, 1).Value2
        free_file = [long]$sheet.Cells.Item(13, 1).Value2
        file_len_after_close = [long]$sheet.Cells.Item(14, 1).Value2
        optional_hash_long = [long]$sheet.Cells.Item(15, 1).Value2
        copied_file_len = [long]$sheet.Cells.Item(16, 1).Value2
        rename_removed_source = [bool]$sheet.Cells.Item(17, 1).Value2
        rename_created_destination = [bool]$sheet.Cells.Item(18, 1).Value2
        hidden_attribute_set = [bool]$sheet.Cells.Item(19, 1).Value2
        killed_file_absent = [bool]$sheet.Cells.Item(20, 1).Value2
        changed_directory = [string]$sheet.Cells.Item(21, 1).Value2
        removed_directory_absent = [bool]$sheet.Cells.Item(22, 1).Value2
        lock_unlock = [string]$sheet.Cells.Item(23, 1).Value2
        width_file_len = [long]$sheet.Cells.Item(24, 1).Value2
        reopen_after_reset_error = [long]$sheet.Cells.Item(25, 1).Value2
        print_bytes_hex = [BitConverter]::ToString(
            [IO.File]::ReadAllBytes((Join-Path $probeRoot 'print.txt'))
        ).Replace('-', '')
        write_bytes_hex = [BitConverter]::ToString(
            [IO.File]::ReadAllBytes((Join-Path $probeRoot 'write.txt'))
        ).Replace('-', '')
        binary_bytes_hex = [BitConverter]::ToString(
            [IO.File]::ReadAllBytes((Join-Path $probeRoot 'binary.dat'))
        ).Replace('-', '')
        optional_hash_bytes_hex = [BitConverter]::ToString(
            [IO.File]::ReadAllBytes((Join-Path $probeRoot 'optional-hash.dat'))
        ).Replace('-', '')
        width_bytes_hex = [BitConverter]::ToString(
            [IO.File]::ReadAllBytes((Join-Path $probeRoot 'width.txt'))
        ).Replace('-', '')
    }
    $expected = [ordered]@{
        line_input = 'alpha 42 omega'
        input_string = 'alpha'
        input_long = 42L
        input_boolean = $true
        multiple_close = 'multi-close-ok'
        binary_lof = 4L
        binary_loc_after_put = 4L
        binary_seek_after_put = 5L
        binary_seek_after_reset = 1L
        binary_get_long = 305419896L
        binary_eof_after_get = $false
        file_len = 0L
        file_len_after_close = 4L
        optional_hash_long = 305419896L
        copied_file_len = 16L
        rename_removed_source = $true
        rename_created_destination = $true
        hidden_attribute_set = $true
        killed_file_absent = $true
        removed_directory_absent = $true
        lock_unlock = 'lock-unlock-ok'
        width_file_len = 17L
        reopen_after_reset_error = 0L
        print_bytes_hex = '616C706861203432206F6D6567610D0A'
        write_bytes_hex = '22616C706861222C34322C2354525545230D0A'
        binary_bytes_hex = '78563412'
        optional_hash_bytes_hex = '78563412'
        width_bytes_hex = '3132333435363738393031323334350D0A'
    }
    foreach ($key in $expected.Keys) {
        if ($result[$key] -ne $expected[$key]) {
            throw "COM result mismatch for ${key}: expected '$($expected[$key])', got '$($result[$key])'"
        }
    }
    if (-not $result.changed_directory.EndsWith('\created-dir', [StringComparison]::OrdinalIgnoreCase)) {
        throw "COM result mismatch for changed_directory: got '$($result.changed_directory)'"
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
    if (-not $KeepArtifacts -and (Test-Path -LiteralPath $probeRoot)) {
        Remove-Item -LiteralPath $probeRoot -Recurse -Force
    }
}
