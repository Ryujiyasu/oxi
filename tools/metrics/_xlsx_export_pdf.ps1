# Prints a workbook to PDF through Excel, for xlsx_pixel_diff.py to compare
# against what Oxi draws.
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
)

$ErrorActionPreference = 'Stop'
# A force-killed Excel leaves a recovery dialog that blocks every later run.
Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery' -Recurse -Force -ErrorAction SilentlyContinue

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open((Resolve-Path $Source).Path)
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
    $wb.ExportAsFixedFormat(0, $Destination)   # 0 = xlTypePDF
    $wb.Close($false)
    Write-Output "ok"
} catch {
    Write-Output ("failed: " + $_.Exception.Message.Split("`n")[0])
} finally {
    $excel.Quit()
}
