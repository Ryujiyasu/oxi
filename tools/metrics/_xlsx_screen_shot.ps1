# Copies a sheet's used range as Excel draws it on screen, for
# xlsx_pixel_diff.py to compare against what Oxi draws.
#
# A printed PDF obeys print areas, page breaks and shrink-to-fit, none of which
# Oxi models, so a workbook whose print area covers one row was being compared
# against the whole sheet. What is on screen is the thing Oxi is trying to draw.
#
# Takes a list of "source<TAB>destination.png" lines, one per workbook.
param([string]$ListFile)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery' -Recurse -Force -ErrorAction SilentlyContinue

$pairs = @()
foreach ($line in Get-Content -LiteralPath $ListFile) {
    if (-not $line.Trim()) { continue }
    $parts = $line -split "`t"
    if ($parts.Count -ge 2) { $pairs += , @($parts[0], $parts[1]) }
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
            $wb = $excel.Workbooks.Open($from, 0, $true)
            $sh = $wb.Worksheets.Item(1)
            # The grid Excel shows behind a sheet is not part of the sheet, and
            # Oxi does not draw it. The sheet has to be the active one before
            # its window will answer, and CopyPicture wants it active anyway.
            $sh.Activate()
            try { $excel.ActiveWindow.DisplayGridlines = $false } catch {}
            # A sheet saved in page-break preview shows a blue frame round its
            # print area, and the picture carries it. That is a property of the
            # window, not of the sheet, so the window is put back to the normal
            # view first. xlNormalView = 1.
            try { $excel.ActiveWindow.View = 1 } catch {}
            # xlScreen = 1, xlBitmap = 2 — what the sheet looks like, not what
            # the printer would be told. The first ask after a workbook opens
            # is often refused and the second one answers, so it is asked again
            # rather than counted as a workbook Excel would not draw.
            for ($try = 1; $try -le 3; $try++) {
                try { $sh.UsedRange.CopyPicture(1, 2); break }
                catch { if ($try -eq 3) { throw }; Start-Sleep -Milliseconds 400 }
            }
            Start-Sleep -Milliseconds 120
            $image = [System.Windows.Forms.Clipboard]::GetImage()
            if ($image) {
                if (Test-Path -LiteralPath $to) { Remove-Item -LiteralPath $to -Force }
                $image.Save($to, [System.Drawing.Imaging.ImageFormat]::Png)
                $image.Dispose()
                Write-Output "ok`t$from"
            } else {
                Write-Output "failed`t$from`tthe clipboard held no picture"
            }
            $wb.Close($false)
        } catch {
            Write-Output ("failed`t$from`t" + $_.Exception.Message.Split("`n")[0])
            foreach ($open in @($excel.Workbooks)) { try { $open.Close($false) } catch {} }
        }
    }
} finally { $excel.Quit() }
