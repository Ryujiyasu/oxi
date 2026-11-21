# Registers oxi-xll-host as the native-messaging host the Oxi extension calls
# (com.oxi.xll_host), for Chrome and Edge, for the current user.
#
#   powershell -File integrations/office-js-extension/native-host/install.ps1 -ExtensionId <id from chrome://extensions>
#
# Build the host first: cd tools/oxi-xll-host; cargo build --release
param(
  [Parameter(Mandatory = $true)][string]$ExtensionId,
  [string]$HostExe = (Resolve-Path "$PSScriptRoot\..\..\..\tools\oxi-xll-host\target\release\oxi-xll-host.exe").Path
)
$manifest = Join-Path $PSScriptRoot 'com.oxi.xll_host.json'
@{
  name = 'com.oxi.xll_host'
  description = 'Oxi XLL host: runs Excel XLL add-ins for the Oxi spreadsheet editor'
  path = $HostExe
  type = 'stdio'
  allowed_origins = @("chrome-extension://$ExtensionId/")
} | ConvertTo-Json | Set-Content -Encoding utf8 $manifest
foreach ($browser in 'Google\Chrome', 'Microsoft\Edge') {
  $key = "HKCU:\Software\$browser\NativeMessagingHosts\com.oxi.xll_host"
  New-Item -Path $key -Force | Out-Null
  Set-ItemProperty -Path $key -Name '(default)' -Value $manifest
}
Write-Output "registered $manifest -> $HostExe"
