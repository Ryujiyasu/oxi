<#
.SYNOPSIS
    Convert test .docx files to PDF using Word COM Automation.
    Run this on a Windows machine with Microsoft Word installed.

.DESCRIPTION
    Opens each .docx in the docx_tests/ folder using Word's COM interface,
    saves as PDF, and closes. This ensures the PDF reflects Word's actual
    text shaping and layout decisions.

.USAGE
    .\word_to_pdf.ps1
    .\word_to_pdf.ps1 -InputDir ".\docx_tests" -OutputDir ".\output\pdfs"
#>

param(
    [string]$InputDir = (Join-Path $PSScriptRoot "docx_tests"),
    [string]$OutputDir = (Join-Path $PSScriptRoot "output\pdfs")
)

# Ensure output directory exists
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

# Start Word
Write-Host "Starting Microsoft Word..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

$docxFiles = Get-ChildItem -Path $InputDir -Filter "*.docx" | Where-Object { $_.Name -ne "~*" }
$total = $docxFiles.Count
$current = 0

Write-Host "Found $total .docx files to convert."

foreach ($file in $docxFiles) {
    $current++
    $pdfName = [System.IO.Path]::ChangeExtension($file.Name, ".pdf")
    $pdfPath = Join-Path $OutputDir $pdfName
    $docxPath = $file.FullName

    Write-Host "[$current/$total] Converting: $($file.Name)"

    try {
        $doc = $word.Documents.Open([string]$docxPath)

        # ExportAsFixedFormat: OutputFileName, ExportFormat (0=PDF), ...
        $doc.ExportAsFixedFormat([string]$pdfPath, 0)
        $doc.Close(0)  # wdDoNotSaveChanges

        Write-Host "         -> $pdfName [OK]"
    }
    catch {
        Write-Host "         -> FAILED: $_" -ForegroundColor Red
    }
}

# Cleanup
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "`nDone. $current PDFs written to $OutputDir"
