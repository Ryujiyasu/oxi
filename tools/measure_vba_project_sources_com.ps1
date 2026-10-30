# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
param([switch]$ExcelOnly)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$fixtureRoot = Join-Path $repoRoot 'tests/fixtures/vba'
$cases = [IO.File]::ReadAllLines((Join-Path $fixtureRoot 'project_sources_ja.tsv'))
$excel = $null
$book = $null
$project = $null
$failures = @()
$compared = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    if ($excel.LanguageSettings.LanguageID(2) -ne 1041) { throw 'This comparison requires a Japanese Excel UI language.' }
    foreach ($file in Get-ChildItem -LiteralPath $fixtureRoot -Filter 'named_project_*.xlsm' -File) {
        $book = $excel.Workbooks.Open($file.FullName, 0, $true)
        $project = $book.VBProject
        if ($null -eq $project) { throw 'Excel VBA project access is disabled. Enable Trust access to the VBA project object model before running this probe.' }
        $projectName = [string]$project.Name
        $helpFile = [string]$excel.Run("'$($book.Name)'!ProjectSourceProbe.VbaDefaultHelpFile")
        foreach ($case in $cases) {
            $first = $case.IndexOf([char]9)
            $second = $case.IndexOf([char]9, $first + 1)
            if ($first -lt 1 -or $second -le $first) { throw "Invalid comparison case: $case" }
            if (-not [string]::Equals($case.Substring(0, $first), $projectName, [StringComparison]::Ordinal)) { continue }
            $procedure = $case.Substring($first + 1, $second - $first - 1)
            $expected = $case.Substring($second + 1).Replace('{VBA_HELP}', $helpFile)
            $actual = [string]$excel.Run("'$($book.Name)'!ProjectSourceProbe.$procedure")
            $compared += 1
            if (-not [string]::Equals($actual, $expected, [StringComparison]::Ordinal)) {
                $failures += [pscustomobject]@{ project = $projectName; procedure = $procedure; expected = $expected; actual = $actual }
            }
        }
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($project)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $project = $null
        $book = $null
    }
    if ($compared -ne $cases.Count) { throw "Compared $compared cases, expected $($cases.Count). Check project names and fixture files." }
    [ordered]@{ version = $excel.Version; build = $excel.Build; total = $compared; passed = $compared - $failures.Count; mismatches = @($failures) } | ConvertTo-Json -Depth 5
    if ($failures.Count -ne 0) { throw 'Excel project Source comparisons differ.' }
}
finally {
    try { if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally {
            foreach ($item in @($project, $book, $excel)) {
                if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) }
            }
        }
    }
}
if (-not $ExcelOnly) {
    & cargo test -p oxivba-core --manifest-path (Join-Path $repoRoot 'Cargo.toml') --test excel_localized_errors renamed_project_sources_match_excel
    if ($LASTEXITCODE -ne 0) { throw "Oxi project Source comparison failed: $LASTEXITCODE" }
}
