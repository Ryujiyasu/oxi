# What Excel makes of a column width: the character count a sheet stores, and
# the pixels the column actually occupies. Oxi has to derive the second from
# the first, and the conversion depends on the standard font.
param([string]$ListFile)

$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    foreach ($line in Get-Content -LiteralPath $ListFile) {
        if (-not $line.Trim()) { continue }
        try {
            $wb = $excel.Workbooks.Open($line.Trim(), 0, $true)
            $sh = $wb.Worksheets.Item(1)
            $font = $wb.Styles.Item('Normal').Font
            $used = $sh.UsedRange
            $first = $used.Column
            $count = [Math]::Min($used.Columns.Count, 6)
            for ($i = 0; $i -lt $count; $i++) {
                $col = $sh.Columns.Item($first + $i)
                # Width is in points; a point is 96/72 pixels on screen.
                Write-Output ("col`t{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f `
                    (Split-Path $line.Trim() -Leaf), ($first + $i), `
                    $col.ColumnWidth, ($col.Width * 96.0 / 72.0), $font.Name, $font.Size)
            }
            $wb.Close($false)
        } catch {
            Write-Output ("failed`t" + $line.Trim() + "`t" + $_.Exception.Message.Split("`n")[0])
            foreach ($open in @($excel.Workbooks)) { try { $open.Close($false) } catch {} }
        }
    }
} finally { $excel.Quit() }
