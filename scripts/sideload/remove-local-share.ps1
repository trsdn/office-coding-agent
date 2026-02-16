param(
  [string]$ShareName = "OfficeAddinCatalog",
  [switch]$RemoveFolder
)

$ErrorActionPreference = 'Stop'

$appRoot = 'HKCU:\Software\office-coding-agent'
$catalogsRoot = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'

$catalogId = $null
if (Test-Path $appRoot) {
  $catalogId = (Get-ItemProperty -Path $appRoot -Name LocalShareCatalogId -ErrorAction SilentlyContinue).LocalShareCatalogId
}

if ($catalogId) {
  $catalogKey = Join-Path $catalogsRoot $catalogId
  if (Test-Path $catalogKey) {
    Remove-Item -Path $catalogKey -Recurse -Force
  }
  Remove-ItemProperty -Path $appRoot -Name LocalShareCatalogId -ErrorAction SilentlyContinue
  Remove-ItemProperty -Path $appRoot -Name LocalShareCatalogUrl -ErrorAction SilentlyContinue
}

$share = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if ($share) {
  $sharePath = $share.Path
  Remove-SmbShare -Name $ShareName -Force
  if ($RemoveFolder -and (Test-Path $sharePath)) {
    Remove-Item -Path $sharePath -Recurse -Force
  }
}

Write-Host "Local share sideload configuration removed."
Write-Host "Restart Excel if it was open."
