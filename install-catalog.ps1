# Installs ppMath as a "shared folder" add-in catalog so PowerPoint keeps it after restart.
# Run from an administrator PowerShell:
#   powershell -ExecutionPolicy Bypass -File .\install-catalog.ps1
#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

$AddinId   = 'eacf1e84-d082-4a1b-bd2d-05400443bc7f'
$CatalogId = '{6f0b6c1e-3a4d-4f3e-9b8a-5d2c1e7a9f10}'
$ShareName = 'ppMathCatalog'
$Dir       = Join-Path $env:USERPROFILE 'ppMathCatalog'

# 1. Folder that holds manifest.xml
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
Copy-Item -Force -Path (Join-Path $PSScriptRoot 'manifest.xml') -Destination $Dir
Write-Host "manifest copied to $Dir"

# 2. Network share of that folder (read-only, current user only)
if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $ShareName -Path $Dir -ReadAccess $env:USERNAME | Out-Null
    Write-Host "share created: \\$env:COMPUTERNAME\$ShareName"
} else {
    Write-Host "share already exists: \\$env:COMPUTERNAME\$ShareName"
}

# 3. Register the share as a trusted add-in catalog (shown in the add-in menu)
$Key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"
New-Item -Force -Path $Key | Out-Null
Set-ItemProperty -Path $Key -Name 'Id'    -Value $CatalogId
Set-ItemProperty -Path $Key -Name 'Url'   -Value "\\$env:COMPUTERNAME\$ShareName"
New-ItemProperty -Path $Key -Name 'Flags' -Value 1 -PropertyType DWord -Force | Out-Null
Write-Host "trusted catalog registered"

# 4. Remove the developer (debug) registration of the same add-in to avoid a duplicate
$Dev = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (Get-ItemProperty -Path $Dev -Name $AddinId -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $Dev -Name $AddinId
    Write-Host "developer registration removed"
}

Write-Host ''
Write-Host 'Done. Restart PowerPoint, then: Home > Add-ins > More Add-ins > SHARED FOLDER > ppMath > Add'
