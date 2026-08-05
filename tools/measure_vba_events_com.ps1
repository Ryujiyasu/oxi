# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$probeRoot = Join-Path ([IO.Path]::GetTempPath()) ("oxivba-events-" + [guid]::NewGuid().ToString('N'))
$workbookPath = Join-Path $probeRoot 'probe.xlsm'
$excel = $null
$workbook = $null
$sheet = $null
$publisher = $null
$listener = $null
$runner = $null

$publisherSource = @'
Option Explicit

Public Event Fired(ByVal Number As Long, ByRef Text As String)
Public Event Ping()

Public Sub Fire()
    Dim message As String
    message = "before"
    RaiseEvent Fired(42, message)
    ThisWorkbook.Worksheets(1).Cells(2, 1).Value2 = message
    RaiseEvent Ping
End Sub
'@

$listenerSource = @'
Option Explicit

Private WithEvents source As OxiPublisher
Private pingCount As Long

Public Sub Attach(ByVal publisher As OxiPublisher)
    Set source = publisher
End Sub

Private Sub source_Fired(ByVal Number As Long, ByRef Text As String)
    ThisWorkbook.Worksheets(1).Cells(1, 1).Value2 = Number
    ThisWorkbook.Worksheets(1).Cells(3, 1).Value2 = Text
    Text = Text & ":handled"
End Sub

Private Sub source_Ping()
    pingCount = pingCount + 1
    ThisWorkbook.Worksheets(1).Cells(4, 1).Value2 = pingCount
End Sub
'@

$runnerSource = @'
Option Explicit

Public Sub OxiEventsProbe()
    Dim publisher As New OxiPublisher
    Dim listener As New OxiListener
    listener.Attach publisher
    publisher.Fire
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
        $publisher = $workbook.VBProject.VBComponents.Add(2)
        $listener = $workbook.VBProject.VBComponents.Add(2)
        $runner = $workbook.VBProject.VBComponents.Add(1)
    }
    catch {
        throw "Excel COM is available, but VBProject access is disabled. $($_.Exception.Message)"
    }
    if ($null -eq $publisher -or $null -eq $listener -or $null -eq $runner) {
        throw 'Excel returned no VBComponent. Enable trust access to the VBA project object model.'
    }

    $publisher.Name = 'OxiPublisher'
    $listener.Name = 'OxiListener'
    $runner.Name = 'OxiEventsModule'
    $publisher.CodeModule.AddFromString($publisherSource)
    $listener.CodeModule.AddFromString($listenerSource)
    $runner.CodeModule.AddFromString($runnerSource)
    $excel.Run("'$($workbook.Name)'!OxiEventsModule.OxiEventsProbe")
    $sheet = $workbook.Worksheets.Item(1)

    $result = [ordered]@{
        delivered_number = [long]$sheet.Cells.Item(1, 1).Value2
        byref_after_handler = [string]$sheet.Cells.Item(2, 1).Value2
        handler_input_text = [string]$sheet.Cells.Item(3, 1).Value2
        zero_arg_event_count = [long]$sheet.Cells.Item(4, 1).Value2
    }
    $expected = [ordered]@{
        delivered_number = 42L
        byref_after_handler = 'before:handled'
        handler_input_text = 'before'
        zero_arg_event_count = 1L
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
    foreach ($comObject in @($sheet, $runner, $listener, $publisher, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    if (Test-Path -LiteralPath $probeRoot) {
        [IO.Directory]::Delete($probeRoot, $true)
    }
}
