# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
param(
    [switch]$ExcelOnly,
    [Alias('Profile')][ValidateSet('portable', 'ja-JP')][string]$ErrorProfile = 'portable'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$fixtureRoot = Join-Path $repoRoot 'tests/fixtures/vba'
$fixture = if ($ErrorProfile -eq 'ja-JP') { 'localized_errors_ja' } else { 'err_properties' }
$source = [IO.File]::ReadAllText((Join-Path $fixtureRoot ($fixture + '.bas')))
$cases = [IO.File]::ReadAllLines((Join-Path $fixtureRoot ($fixture + '.tsv')))
$excel = $null
$book = $null
$component = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $project = $book.VBProject
    if ($null -eq $project) {
        throw 'Excel VBA project access is disabled. Enable Trust access to the VBA project object model before running this probe.'
    }
    $component = $project.VBComponents.Add(1)
    $component.Name = 'ErrCompatibility'
    $component.CodeModule.AddFromString($source)
    $helpFile = ''
    if ($ErrorProfile -eq 'ja-JP') {
        if ($excel.LanguageSettings.LanguageID(2) -ne 1041) {
            throw 'The ja-JP comparison requires a Japanese Excel UI language.'
        }
        $helpFile = [string]$excel.Run("'$($book.Name)'!ErrCompatibility.VbaDefaultHelpFile")
        if ([string]::IsNullOrEmpty($helpFile)) { throw 'Excel did not supply its default VBA help path.' }
    }
    $failures = @()
    foreach ($case in $cases) {
        $separator = $case.IndexOf([char]9)
        if ($separator -lt 1) { throw "Invalid comparison case: $case" }
        $procedureName = $case.Substring(0, $separator)
        $actual = [string]$excel.Run("'$($book.Name)'!ErrCompatibility.$procedureName")
        $expected = $case.Substring($separator + 1).Replace('{VBA_HELP}', $helpFile)
        if (-not [string]::Equals($actual, $expected, [StringComparison]::Ordinal)) {
            $failures += [pscustomobject]@{ case = $procedureName; expected = $expected; actual = $actual }
        }
    }
    [ordered]@{
        version = $excel.Version
        build = $excel.Build
        operating_system = $excel.OperatingSystem
        profile = $ErrorProfile
        total = $cases.Count
        passed = $cases.Count - $failures.Count
        mismatches = @($failures)
    } | ConvertTo-Json -Depth 5
    if ($failures.Count -ne 0) { throw "$($failures.Count) Excel result(s) differ from the comparison fixture." }
}
finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
    }
    finally {
        try {
            if ($null -ne $excel) { $excel.Quit() }
        }
        finally {
            foreach ($item in @($component, $project, $book, $excel)) {
                if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)
                }
            }
        }
    }
}
if (-not $ExcelOnly) {
    $testName = if ($ErrorProfile -eq 'ja-JP') { 'excel_localized_errors' } else { 'excel_err_properties' }
    & cargo test -p oxivba-core --manifest-path (Join-Path $repoRoot 'Cargo.toml') --test $testName
    if ($LASTEXITCODE -ne 0) { throw "Oxi comparison failed with exit code $LASTEXITCODE" }
}
