# Removes what install-catalog.ps1 created.
# Run from an administrator PowerShell:
#   powershell -ExecutionPolicy Bypass -File .\uninstall-catalog.ps1
#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

$CatalogId = '{6f0b6c1e-3a4d-4f3e-9b8a-5d2c1e7a9f10}'
$ShareName = 'ppMathCatalog'
$Dir       = Join-Path $env:USERPROFILE 'ppMathCatalog'

$Key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"
if (Test-Path $Key) { Remove-Item -Path $Key -Recurse; Write-Host 'trusted catalog removed' }
if (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue) {
    Remove-SmbShare -Name $ShareName -Force; Write-Host 'share removed'
}
if (Test-Path $Dir) { Remove-Item -Path $Dir -Recurse -Force; Write-Host 'folder removed' }
