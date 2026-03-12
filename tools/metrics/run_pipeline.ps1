<#
.SYNOPSIS
    Full metrics pipeline: generate docx → Word PDF → extract & compare.
    Run this on Windows with Word installed.

.USAGE
    .\run_pipeline.ps1
#>

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== Oxidocs Metrics Pipeline ===" -ForegroundColor Cyan
Write-Host ""

# Step 0: Install Python dependencies
Write-Host "[0/3] Installing Python dependencies..." -ForegroundColor Yellow
pip install -r (Join-Path $scriptDir "requirements.txt") --quiet
Write-Host ""

# Step 1: Generate test docx files
Write-Host "[1/3] Generating test .docx files..." -ForegroundColor Yellow
python (Join-Path $scriptDir "generate_test_docx.py")
Write-Host ""

# Step 2: Convert to PDF via Word COM
Write-Host "[2/3] Converting to PDF via Word COM Automation..." -ForegroundColor Yellow
& (Join-Path $scriptDir "word_to_pdf.ps1")
Write-Host ""

# Step 3: Extract metrics and compare
Write-Host "[3/3] Extracting metrics and comparing..." -ForegroundColor Yellow
python (Join-Path $scriptDir "extract_metrics.py")
Write-Host ""

Write-Host "=== Pipeline complete ===" -ForegroundColor Green
Write-Host "Results: $(Join-Path $scriptDir 'output\metrics_diff.json')"
