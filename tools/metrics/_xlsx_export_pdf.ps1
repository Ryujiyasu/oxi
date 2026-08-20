# Prints workbooks to PDF through Excel, for xlsx_pixel_diff.py to compare
# against what Oxi draws.
#
# Takes either one workbook (-Source/-Destination) or a list of them
# (-ListFile, one "source<TAB>destination" per line). The list keeps Excel open
# across the whole batch, which is the difference between minutes and an hour
# over a few hundred files.
param(
    [string]$Source,
    [string]$Destination,
    [string]$ListFile
)

$ErrorActionPreference = 'Stop'
# A force-killed Excel leaves a recovery dialog that blocks every later run.
Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery' -Recurse -Force -ErrorAction SilentlyContinue

$pairs = @()
if ($ListFile) {
    foreach ($line in Get-Content -LiteralPath $ListFile) {
        if (-not $line.Trim()) { continue }
        $parts = $line -split "`t"
        if ($parts.Count -ge 2) { $pairs += , @($parts[0], $parts[1]) }
    }
} elseif ($Source -and $Destination) {
    $pairs += , @($Source, $Destination)
} else {
    Write-Output 'failed: give -Source and -Destination, or -ListFile'
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    foreach ($pair in $pairs) {
        $from = $pair[0]
        $to = $pair[1]
        try {
            $wb = $excel.Workbooks.Open($from, 0, $true)   # read-only, no link update
            if (Test-Path -LiteralPath $to) { Remove-Item -LiteralPath $to -Force }
            $wb.ExportAsFixedFormat(0, $to)   # 0 = xlTypePDF
            $wb.Close($false)
            Write-Output "ok`t$from"
        } catch {
            Write-Output ("failed`t$from`t" + $_.Exception.Message.Split("`n")[0])
            foreach ($open in @($excel.Workbooks)) {
                try { $open.Close($false) } catch {}
            }
        }
    }
} finally {
    $excel.Quit()
}
