param(
  [string]$ShareName = "OfficeAddinCatalog",
  [string]$SharePath = "$env:USERPROFILE\OfficeAddinCatalog"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $SharePath)) {
  New-Item -ItemType Directory -Path $SharePath -Force | Out-Null
}

$existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if (-not $existing) {
  New-SmbShare -Name $ShareName -Path $SharePath -FullAccess "$env:USERNAME" -ErrorAction Stop | Out-Null
}

$uncPath = "\\$env:COMPUTERNAME\$ShareName"

Write-Host "Local catalog share is ready."
Write-Host "Share path: $SharePath"
Write-Host "UNC path:   $uncPath"
Write-Host ""
Write-Host "Next:"
Write-Host "  1) npm run sideload:share:trust"
Write-Host "  2) npm run manifest:staging"
Write-Host "  3) npm run sideload:share:publish"
