# Registers oxi-xll-host and oxi-com-host as the native-messaging hosts the
# Oxi extension calls (com.oxi.xll_host, com.oxi.com_host), for Chrome and
# Edge, for the current user.
#
#   powershell -File integrations/office-js-extension/native-host/install.ps1 -ExtensionId <id from chrome://extensions>
#
# Build the hosts first: cargo build --release in tools/oxi-xll-host and
# tools/oxi-com-host. A host that is not built is skipped.
param(
  [Parameter(Mandatory = $true)][string]$ExtensionId
)
$tools = Resolve-Path "$PSScriptRoot\..\..\..\tools"
$hosts = @(
  @{ name = 'com.oxi.xll_host'; exe = "$tools\oxi-xll-host\target\release\oxi-xll-host.exe"; what = 'Oxi XLL host: runs Excel XLL add-ins for the Oxi spreadsheet editor' },
  @{ name = 'com.oxi.com_host'; exe = "$tools\oxi-com-host\target\release\oxi-com-host.exe"; what = 'Oxi COM host: calls Excel COM Automation add-ins for the Oxi spreadsheet editor' }
)
foreach ($h in $hosts) {
  if (-not (Test-Path $h.exe)) { Write-Output "skipped $($h.name): $($h.exe) is not built"; continue }
  $manifest = Join-Path $PSScriptRoot "$($h.name).json"
  @{
    name = $h.name
    description = $h.what
    path = $h.exe
    type = 'stdio'
    allowed_origins = @("chrome-extension://$ExtensionId/")
  } | ConvertTo-Json | Set-Content -Encoding utf8 $manifest
  foreach ($browser in 'Google\Chrome', 'Microsoft\Edge') {
    $key = "HKCU:\Software\$browser\NativeMessagingHosts\$($h.name)"
    New-Item -Path $key -Force | Out-Null
    Set-ItemProperty -Path $key -Name '(default)' -Value $manifest
  }
  Write-Output "registered $manifest -> $($h.exe)"
}
